VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{A16AEA80-F8AA-4014-8A42-5B898C073696}#13.0#0"; "QMLISTBOX.OCX"
Begin VB.Form frmma 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H0047491F&
   BorderStyle     =   0  'None
   Caption         =   "iCee"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5130
   DrawWidth       =   10
   FillStyle       =   0  'Solid
   Icon            =   "主窗体.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   342
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PF 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00423E39&
      BorderStyle     =   0  'None
      Height          =   9105
      Index           =   1
      Left            =   4200
      ScaleHeight     =   607
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   125
      Top             =   9360
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.PictureBox iFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9360
      IMEMode         =   1  'ON
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   624
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   340
      TabIndex        =   0
      Top             =   240
      Width           =   5100
      Begin VB.PictureBox PICDL 
         AutoRedraw      =   -1  'True
         BackColor       =   &H005B1FC0&
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   3840
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   72
         TabIndex        =   210
         Top             =   1680
         Width           =   1080
         Begin VB.Shape SB 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   0  'Transparent
            Height          =   45
            Index           =   1
            Left            =   0
            Top             =   570
            Width           =   15
         End
         Begin VB.Label LA 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   19
            Left            =   720
            TabIndex        =   211
            Top             =   150
            Width           =   165
         End
      End
      Begin VB.PictureBox PF 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00201400&
         BorderStyle     =   0  'None
         Height          =   9375
         Index           =   12
         Left            =   3960
         ScaleHeight     =   625
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   340
         TabIndex        =   59
         Top             =   8520
         Visible         =   0   'False
         Width           =   5100
         Begin VB.Timer TMTIM 
            Interval        =   1000
            Left            =   3600
            Top             =   600
         End
         Begin VB.PictureBox PTIME 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1170
            Index           =   4
            Left            =   3840
            ScaleHeight     =   78
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   54
            TabIndex        =   227
            Top             =   2280
            Width           =   810
         End
         Begin VB.PictureBox PTIME 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1170
            Index           =   3
            Left            =   3000
            ScaleHeight     =   78
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   54
            TabIndex        =   226
            Top             =   2280
            Width           =   810
         End
         Begin VB.PictureBox PTIME 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1170
            Index           =   2
            Left            =   2160
            ScaleHeight     =   78
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   54
            TabIndex        =   225
            Top             =   2280
            Width           =   810
         End
         Begin VB.PictureBox PTIME 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1170
            Index           =   1
            Left            =   1320
            ScaleHeight     =   78
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   54
            TabIndex        =   224
            Top             =   2280
            Width           =   810
         End
         Begin VB.PictureBox PTIME 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1170
            Index           =   0
            Left            =   480
            ScaleHeight     =   78
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   54
            TabIndex        =   223
            Top             =   2280
            Width           =   810
         End
         Begin ICEE.ICEE_KEY ICL 
            Height          =   525
            Index           =   0
            Left            =   840
            TabIndex        =   86
            Top             =   6600
            Width           =   1755
            _ExtentX        =   4260
            _ExtentY        =   926
         End
         Begin VB.TextBox TXTPOUP 
            Appearance      =   0  'Flat
            BackColor       =   &H00AEC3D2&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   1440
            MaxLength       =   16
            TabIndex        =   60
            Top             =   6120
            Width           =   2295
         End
         Begin ICEE.ICEE_KEY ICL 
            Height          =   525
            Index           =   17
            Left            =   2760
            TabIndex        =   228
            Top             =   6600
            Width           =   1755
            _ExtentX        =   4260
            _ExtentY        =   926
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Index           =   30
            Left            =   3000
            TabIndex        =   229
            Top             =   4920
            Width           =   150
         End
         Begin VB.Shape SB 
            BackColor       =   &H0000B9FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   0  'Transparent
            Height          =   45
            Index           =   2
            Left            =   0
            Top             =   4050
            Width           =   15
         End
      End
      Begin VB.PictureBox PP 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H002A1C05&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6600
         Left            =   840
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   440
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   340
         TabIndex        =   4
         Top             =   2640
         Visible         =   0   'False
         Width           =   5100
         Begin VB.PictureBox Pmusic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00EFBC44&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   5325
            Left            =   0
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   355
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   336
            TabIndex        =   5
            Top             =   5280
            Width           =   5040
            Begin VB.PictureBox IU 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H005B1FC0&
               BorderStyle     =   0  'None
               Height          =   600
               Left            =   0
               OLEDropMode     =   1  'Manual
               ScaleHeight     =   40
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   40
               TabIndex        =   314
               Top             =   0
               Width           =   600
            End
            Begin VB.PictureBox Mbar 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H007A7417&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   405
               Left            =   0
               ScaleHeight     =   27
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   335
               TabIndex        =   6
               Top             =   4560
               Width           =   5025
               Begin ICEE.ICEE_KEY ICP 
                  Height          =   375
                  Index           =   0
                  Left            =   240
                  TabIndex        =   181
                  Top             =   0
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   661
               End
               Begin ICEE.ICEE_KEY ICP 
                  Height          =   375
                  Index           =   1
                  Left            =   1440
                  TabIndex        =   182
                  Top             =   0
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   661
               End
               Begin ICEE.ICEE_KEY ICP 
                  Height          =   375
                  Index           =   2
                  Left            =   2640
                  TabIndex        =   183
                  Top             =   0
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   661
               End
               Begin ICEE.ICEE_KEY ICP 
                  Height          =   375
                  Index           =   3
                  Left            =   3840
                  TabIndex        =   184
                  Top             =   0
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   661
               End
            End
            Begin VB.PictureBox Pser 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00828637&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   450
               Left            =   0
               ScaleHeight     =   30
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   342
               TabIndex        =   8
               Top             =   4080
               Visible         =   0   'False
               Width           =   5130
               Begin VB.TextBox TTS 
                  BackColor       =   &H00828637&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H00FFFFFF&
                  Height          =   225
                  IMEMode         =   1  'ON
                  Left            =   120
                  TabIndex        =   9
                  Text            =   "快速定位列表内歌曲"
                  Top             =   120
                  Width           =   4695
               End
            End
            Begin MusicListBox.QMListBox PLIST 
               Height          =   2295
               Left            =   0
               TabIndex        =   127
               Top             =   600
               Width           =   4800
               _ExtentX        =   8467
               _ExtentY        =   4048
               BackColor       =   349412
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ItemHeight      =   40
               ItemBkColor1    =   1533689
               ItemBkColor2    =   349412
               ItemSelBkColor  =   9151263
               ItemNorTextColor=   14737632
               ItemSelTextColor=   16777215
               ItemPlyTextColor=   16777215
               Border          =   0   'False
               BorderColor     =   12821368
               BorderWidth     =   1
            End
            Begin VB.Label LA 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "New"
               BeginProperty Font 
                  Name            =   "微软雅黑"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   18
               Left            =   720
               TabIndex        =   315
               Top             =   210
               Width           =   480
            End
            Begin WMPLibCtl.WindowsMediaPlayer Wm 
               Height          =   1215
               Left            =   2880
               TabIndex        =   17
               Top             =   14985
               Visible         =   0   'False
               Width           =   1530
               URL             =   ""
               rate            =   1
               balance         =   0
               currentPosition =   0
               defaultFrame    =   ""
               playCount       =   1
               autoStart       =   -1  'True
               currentMarker   =   0
               invokeURLs      =   0   'False
               baseURL         =   ""
               volume          =   100
               mute            =   0   'False
               uiMode          =   "none"
               stretchToFit    =   0   'False
               windowlessVideo =   0   'False
               enabled         =   0   'False
               enableContextMenu=   -1  'True
               fullScreen      =   0   'False
               SAMIStyle       =   ""
               SAMILang        =   ""
               SAMIFilename    =   ""
               captioningID    =   ""
               enableErrorDialogs=   0   'False
               _cx             =   2699
               _cy             =   2143
            End
         End
         Begin VB.PictureBox PF 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H008080FF&
            BorderStyle     =   0  'None
            Height          =   3735
            Index           =   13
            Left            =   0
            ScaleHeight     =   249
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   333
            TabIndex        =   293
            Top             =   1560
            Visible         =   0   'False
            Width           =   4995
            Begin VB.PictureBox IMCLEAR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   375
               Left            =   120
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   80
               TabIndex        =   313
               Top             =   180
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.PictureBox PMINFO 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H006D5818&
               BorderStyle     =   0  'None
               Height          =   375
               Left            =   120
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   80
               TabIndex        =   312
               Top             =   180
               Visible         =   0   'False
               Width           =   1200
            End
            Begin ICEE.ICEE_WIN8 PMU 
               Height          =   990
               Index           =   0
               Left            =   120
               TabIndex        =   300
               Top             =   960
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   1746
            End
            Begin VB.PictureBox PCOO 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H000080FF&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   4
               Left            =   2040
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   298
               Top             =   3195
               Width           =   375
            End
            Begin VB.PictureBox PCOO 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   3
               Left            =   600
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   297
               Top             =   3195
               Width           =   375
            End
            Begin VB.PictureBox PCOO 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00DB59D8&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   2
               Left            =   1560
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   296
               Top             =   3195
               Width           =   375
            End
            Begin VB.PictureBox PCOO 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H002EBC7C&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   1
               Left            =   1080
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   295
               Top             =   3195
               Width           =   375
            End
            Begin VB.PictureBox PCOO 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00AA7402&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   0
               Left            =   120
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   294
               Top             =   3195
               Width           =   375
            End
            Begin ICEE.ICEE_WIN8 PMU 
               Height          =   990
               Index           =   1
               Left            =   1200
               TabIndex        =   301
               Top             =   960
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   1746
            End
            Begin ICEE.ICEE_WIN8 PMU 
               Height          =   990
               Index           =   3
               Left            =   2280
               TabIndex        =   302
               Top             =   960
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   1746
            End
            Begin ICEE.ICEE_WIN8 PMU 
               Height          =   990
               Index           =   4
               Left            =   1200
               TabIndex        =   303
               Top             =   2040
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   1746
            End
            Begin ICEE.ICEE_WIN8 PMU 
               Height          =   990
               Index           =   2
               Left            =   120
               TabIndex        =   304
               Top             =   2040
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   1746
            End
         End
         Begin VB.PictureBox IMSERG 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H006D5818&
            BorderStyle     =   0  'None
            Height          =   720
            Left            =   4350
            ScaleHeight     =   48
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   48
            TabIndex        =   291
            Top             =   1560
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.PictureBox PKU 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H006D5818&
            BorderStyle     =   0  'None
            Height          =   720
            Left            =   4320
            ScaleHeight     =   48
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   48
            TabIndex        =   289
            Top             =   720
            Width           =   720
         End
         Begin VB.PictureBox PICBACK 
            AutoRedraw      =   -1  'True
            BackColor       =   &H005B1FC0&
            BorderStyle     =   0  'None
            Height          =   600
            Left            =   0
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   40
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   88
            TabIndex        =   62
            Top             =   0
            Width           =   1320
            Begin VB.Label LA 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "999"
               BeginProperty Font 
                  Name            =   "微软雅黑"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   20
               Left            =   720
               TabIndex        =   123
               Top             =   210
               Width           =   435
            End
         End
         Begin VB.PictureBox PV 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H006D5818&
            BorderStyle     =   0  'None
            Height          =   450
            Left            =   120
            ScaleHeight     =   450
            ScaleWidth      =   4830
            TabIndex        =   10
            Top             =   5760
            Visible         =   0   'False
            Width           =   4830
            Begin VB.Shape ML 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H0028D985&
               BorderStyle     =   0  'Transparent
               FillColor       =   &H0028D985&
               Height          =   75
               Left            =   0
               Top             =   240
               Width           =   4335
            End
         End
         Begin VB.PictureBox PZOR 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H005757EE&
            BorderStyle     =   0  'None
            Height          =   450
            Left            =   120
            ScaleHeight     =   30
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   175
            Top             =   5760
            Width           =   525
         End
         Begin VB.PictureBox IPRE 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H005757EE&
            BorderStyle     =   0  'None
            Height          =   690
            Left            =   840
            ScaleHeight     =   46
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   46
            TabIndex        =   174
            Top             =   5640
            Width           =   690
         End
         Begin VB.PictureBox IMVOL 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H005757EE&
            BorderStyle     =   0  'None
            Height          =   450
            Left            =   4500
            ScaleHeight     =   30
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   173
            Top             =   5760
            Width           =   525
         End
         Begin VB.PictureBox INEXT 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H005757EE&
            BorderStyle     =   0  'None
            Height          =   690
            Left            =   3480
            ScaleHeight     =   46
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   46
            TabIndex        =   172
            Top             =   5640
            Width           =   690
         End
         Begin VB.PictureBox PLAYB 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H005757EE&
            BorderStyle     =   0  'None
            Height          =   1260
            Left            =   1920
            ScaleHeight     =   84
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   84
            TabIndex        =   171
            Top             =   5280
            Width           =   1260
         End
         Begin VB.PictureBox PMDL 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H006D5818&
            BorderStyle     =   0  'None
            Height          =   720
            Left            =   120
            ScaleHeight     =   48
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   48
            TabIndex        =   124
            Top             =   1560
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.PictureBox PSEND 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H006D5818&
            BorderStyle     =   0  'None
            Height          =   720
            Left            =   120
            ScaleHeight     =   48
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   48
            TabIndex        =   99
            Top             =   1560
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.PictureBox ISHA 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   600
            Left            =   4485
            ScaleHeight     =   40
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   40
            TabIndex        =   299
            Top             =   3840
            Width           =   600
         End
         Begin VB.Label LBSINGER 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "未知歌手"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   600
            TabIndex        =   290
            Top             =   1170
            Width           =   720
         End
         Begin VB.Image IMGFAV 
            Height          =   240
            Left            =   120
            Picture         =   "主窗体.frx":0FA2
            Top             =   960
            Width           =   240
         End
         Begin VB.Label LA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "未知"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   10
            Left            =   600
            TabIndex        =   92
            Top             =   4800
            Width           =   3600
         End
         Begin VB.Label LA 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "专辑:"
            ForeColor       =   &H00E0E0E0&
            Height          =   180
            Index           =   8
            Left            =   120
            TabIndex        =   91
            Top             =   4800
            Width           =   450
         End
         Begin VB.Label LA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "未知"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   7
            Left            =   600
            TabIndex        =   90
            Top             =   5040
            Width           =   3690
         End
         Begin VB.Label LA 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "年代:"
            ForeColor       =   &H00E0E0E0&
            Height          =   180
            Index           =   6
            Left            =   120
            TabIndex        =   89
            Top             =   5040
            Width           =   450
         End
         Begin VB.Image EI 
            Enabled         =   0   'False
            Height          =   30
            Left            =   240
            OLEDropMode     =   1  'Manual
            Stretch         =   -1  'True
            Top             =   6585
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label LBSONG 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "HI,我是歌曲名称"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   600
            OLEDropMode     =   1  'Manual
            TabIndex        =   20
            Top             =   840
            Width           =   1365
         End
         Begin VB.Image K 
            Height          =   30
            Index           =   5
            Left            =   195
            Stretch         =   -1  'True
            Top             =   6585
            Visible         =   0   'False
            Width           =   4725
         End
      End
      Begin VB.PictureBox PF 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00383636&
         BorderStyle     =   0  'None
         Height          =   7125
         Index           =   15
         Left            =   4800
         ScaleHeight     =   475
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   340
         TabIndex        =   87
         Top             =   3840
         Visible         =   0   'False
         Width           =   5100
         Begin VB.PictureBox PF 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   2535
            Index           =   14
            Left            =   420
            ScaleHeight     =   169
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   293
            TabIndex        =   305
            Top             =   1680
            Width           =   4395
            Begin ICEE.IVScroll SURO 
               Height          =   2415
               Left            =   4080
               TabIndex        =   307
               Top             =   75
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   4260
               MinV            =   0
               MaxV            =   20
               Value           =   0
               SmallChange     =   1
               LargeChange     =   10
            End
            Begin VB.PictureBox PF 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   7335
               Index           =   16
               Left            =   120
               ScaleHeight     =   489
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   257
               TabIndex        =   306
               Top             =   240
               Width           =   3855
               Begin VB.Image IMSIGN 
                  Height          =   360
                  Left            =   360
                  Picture         =   "主窗体.frx":132C
                  Top             =   0
                  Width           =   360
               End
               Begin VB.Image MNO 
                  Height          =   1245
                  Left            =   720
                  Picture         =   "主窗体.frx":1A96
                  Stretch         =   -1  'True
                  Top             =   6000
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   25
                  Left            =   0
                  Stretch         =   -1  'True
                  Top             =   6000
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   24
                  Left            =   2880
                  Stretch         =   -1  'True
                  Top             =   4800
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   23
                  Left            =   2160
                  Stretch         =   -1  'True
                  Top             =   4800
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   22
                  Left            =   1440
                  Stretch         =   -1  'True
                  Top             =   4800
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   21
                  Left            =   720
                  Stretch         =   -1  'True
                  Top             =   4800
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   20
                  Left            =   0
                  Stretch         =   -1  'True
                  Top             =   4800
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   19
                  Left            =   0
                  Stretch         =   -1  'True
                  Top             =   3600
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   18
                  Left            =   2160
                  Stretch         =   -1  'True
                  Top             =   3600
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   17
                  Left            =   1440
                  Stretch         =   -1  'True
                  Top             =   3600
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   16
                  Left            =   720
                  Stretch         =   -1  'True
                  Top             =   3600
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   15
                  Left            =   2880
                  Stretch         =   -1  'True
                  Top             =   3600
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   14
                  Left            =   2880
                  Stretch         =   -1  'True
                  Top             =   2400
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   13
                  Left            =   2160
                  Stretch         =   -1  'True
                  Top             =   2400
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   12
                  Left            =   1440
                  Stretch         =   -1  'True
                  Top             =   2400
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   11
                  Left            =   720
                  Stretch         =   -1  'True
                  Top             =   2400
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   10
                  Left            =   0
                  Stretch         =   -1  'True
                  Top             =   2400
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   9
                  Left            =   2160
                  Stretch         =   -1  'True
                  Top             =   1200
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   8
                  Left            =   1440
                  Stretch         =   -1  'True
                  Top             =   1200
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   7
                  Left            =   720
                  Stretch         =   -1  'True
                  Top             =   1200
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   5
                  Left            =   2880
                  Stretch         =   -1  'True
                  Top             =   1200
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   4
                  Left            =   2880
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   3
                  Left            =   2160
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   2
                  Left            =   1440
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   1
                  Left            =   720
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   0
                  Left            =   0
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   750
               End
               Begin VB.Image MBK 
                  Height          =   1245
                  Index           =   6
                  Left            =   0
                  Stretch         =   -1  'True
                  Top             =   1200
                  Width           =   750
               End
            End
         End
         Begin ICEE.ICEE_KEY IST 
            Height          =   1095
            Index           =   0
            Left            =   480
            TabIndex        =   177
            Top             =   480
            Width           =   1095
            _ExtentX        =   3413
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY IST 
            Height          =   1095
            Index           =   1
            Left            =   1560
            TabIndex        =   178
            Top             =   480
            Width           =   1095
            _ExtentX        =   3413
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY IST 
            Height          =   1095
            Index           =   2
            Left            =   2640
            TabIndex        =   179
            Top             =   480
            Width           =   1095
            _ExtentX        =   3413
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY IST 
            Height          =   1095
            Index           =   3
            Left            =   3720
            TabIndex        =   180
            Top             =   480
            Width           =   1095
            _ExtentX        =   3413
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICS 
            Height          =   495
            Index           =   0
            Left            =   480
            TabIndex        =   194
            Top             =   4320
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICS 
            Height          =   495
            Index           =   1
            Left            =   1680
            TabIndex        =   195
            Top             =   4320
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
         End
         Begin VB.PictureBox PF 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H002DA6FF&
            BorderStyle     =   0  'None
            Height          =   1335
            Index           =   5
            Left            =   480
            ScaleHeight     =   89
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   281
            TabIndex        =   142
            Top             =   4800
            Width           =   4215
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00DAA52D&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   17
               Left            =   3000
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   255
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H0027FEC2&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   6
               Left            =   3360
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   149
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H0001A175&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   2
               Left            =   2640
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   145
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00C48356&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   1
               Left            =   3000
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   144
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H0062B728&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   31
               Left            =   3000
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   268
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00CCD800&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   9
               Left            =   2280
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   152
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00B0B024&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   27
               Left            =   2280
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   264
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H0092DC61&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   24
               Left            =   2280
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   261
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H001611EA&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   7
               Left            =   1560
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   150
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H001414CD&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   28
               Left            =   1560
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   265
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00686815&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   16
               Left            =   1560
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   254
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H000554E4&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   3
               Left            =   840
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   146
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H005273E8&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   30
               Left            =   840
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   267
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H0000035E&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   18
               Left            =   840
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   256
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   10
               Left            =   120
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   153
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00808080&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   13
               Left            =   120
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   251
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H008F61EA&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   22
               Left            =   120
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   260
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00414141&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   12
               Left            =   480
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   176
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H0000B9FF&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   11
               Left            =   1920
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   154
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H0084536F&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   8
               Left            =   3360
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   151
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H006623D6&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   4
               Left            =   1200
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   147
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00F4C931&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   0
               Left            =   2640
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   143
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00565656&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   20
               Left            =   480
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   258
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H000082E1&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   26
               Left            =   1920
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   263
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00521FA7&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   29
               Left            =   1200
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   266
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H0000DAF9&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   32
               Left            =   2640
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   269
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H007E5502&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   14
               Left            =   1200
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   252
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H006826D5&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   21
               Left            =   480
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   259
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00009CB3&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   25
               Left            =   1920
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   262
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00CF5F38&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   19
               Left            =   3360
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   257
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00563F30&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   15
               Left            =   3720
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   253
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00606015&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   5
               Left            =   3720
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   148
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox CZ_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00261700&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   33
               Left            =   3720
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   279
               Top             =   120
               Width           =   375
            End
         End
         Begin VB.PictureBox PF 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H002DA6FF&
            BorderStyle     =   0  'None
            Height          =   1335
            Index           =   7
            Left            =   480
            ScaleHeight     =   89
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   281
            TabIndex        =   196
            Top             =   4800
            Visible         =   0   'False
            Width           =   4215
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00565656&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   11
               Left            =   3720
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   208
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H0000035E&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   14
               Left            =   3000
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   232
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H004743F8&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   21
               Left            =   2280
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   239
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00B9C127&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   24
               Left            =   1560
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   242
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00606015&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   5
               Left            =   840
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   202
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H0023E7B0&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   4
               Left            =   120
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   205
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H001611EA&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   10
               Left            =   1200
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   200
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H0001A175&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   6
               Left            =   480
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   201
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H000DECC5&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   20
               Left            =   1920
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   238
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H0058541F&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   19
               Left            =   1560
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   237
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00E9C07B&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   16
               Left            =   3720
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   234
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H0084536F&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   15
               Left            =   3000
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   233
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H008F61EA&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   22
               Left            =   1920
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   240
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00ACB500&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   18
               Left            =   2280
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   236
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00E59147&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   2
               Left            =   840
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   206
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00CCD800&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   0
               Left            =   120
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   198
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00BB2F47&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   3
               Left            =   1200
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   199
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00F4C931&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   1
               Left            =   480
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   207
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00404040&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   28
               Left            =   3000
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   246
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00261700&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   9
               Left            =   2280
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   203
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00945AE2&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   29
               Left            =   1200
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   247
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H009F5FC7&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   26
               Left            =   1920
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   244
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H0000B9FF&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   7
               Left            =   480
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   197
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H0000DCFF&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   12
               Left            =   120
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   209
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H000554E4&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   8
               Left            =   840
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   204
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H008047DE&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   30
               Left            =   1560
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   248
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H0092DC61&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   27
               Left            =   3720
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   245
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00767618&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   13
               Left            =   3360
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   231
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H006826D5&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   32
               Left            =   3360
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   250
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00338857&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   25
               Left            =   3360
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   243
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   31
               Left            =   2640
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   249
               Top             =   840
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00C48356&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   23
               Left            =   2640
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   241
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox JM_COLOR 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H0027FEC2&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   17
               Left            =   2640
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   235
               Top             =   480
               Width           =   375
            End
         End
      End
      Begin VB.PictureBox PF 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         BorderStyle     =   0  'None
         Height          =   5415
         Index           =   0
         Left            =   3960
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   361
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   340
         TabIndex        =   63
         Top             =   6840
         Visible         =   0   'False
         Width           =   5100
         Begin ICEE.ICEE_KEY ICL 
            Height          =   495
            Index           =   9
            Left            =   690
            TabIndex        =   185
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICL 
            Height          =   495
            Index           =   10
            Left            =   2880
            TabIndex        =   186
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
         End
         Begin VB.TextBox txtEntry 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   720
            TabIndex        =   64
            Top             =   1440
            Width           =   3615
         End
         Begin VB.PictureBox PF 
            BackColor       =   &H00231C09&
            BorderStyle     =   0  'None
            Height          =   2325
            Index           =   9
            Left            =   675
            ScaleHeight     =   155
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   247
            TabIndex        =   66
            Top             =   2415
            Visible         =   0   'False
            Width           =   3705
            Begin VB.TextBox txtLogBase 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H001F1FE2&
               Height          =   285
               Left            =   840
               TabIndex        =   82
               Text            =   "10"
               Top             =   1800
               Width           =   612
            End
            Begin VB.TextBox txtDecimal 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   73
               Text            =   "9"
               Top             =   1320
               Width           =   255
            End
            Begin VB.VScrollBar vsbDecimal 
               Height          =   285
               Left            =   1680
               Max             =   10
               TabIndex        =   72
               Top             =   1335
               Width           =   255
            End
            Begin VB.OptionButton optBaseMode 
               BackColor       =   &H00231C09&
               Caption         =   "十六进制"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   71
               Top             =   960
               Width           =   1140
            End
            Begin VB.OptionButton optBaseMode 
               BackColor       =   &H00231C09&
               Caption         =   "八进制"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   3
               Left            =   2040
               TabIndex        =   70
               Top             =   600
               Width           =   900
            End
            Begin VB.OptionButton optBaseMode 
               BackColor       =   &H00231C09&
               Caption         =   "二进制"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   69
               Top             =   600
               Width           =   900
            End
            Begin VB.OptionButton optBaseMode 
               BackColor       =   &H00231C09&
               Caption         =   "十进制"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   1080
               TabIndex        =   68
               Top             =   600
               Value           =   -1  'True
               Width           =   900
            End
            Begin VB.Frame fraAngleM 
               Appearance      =   0  'Flat
               BackColor       =   &H00231C09&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   0
               TabIndex        =   67
               Top             =   120
               Width           =   1815
               Begin VB.OptionButton optAngMode 
                  BackColor       =   &H00231C09&
                  Caption         =   "弧度"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Index           =   1
                  Left            =   960
                  TabIndex        =   85
                  Top             =   0
                  Width           =   780
               End
               Begin VB.OptionButton optAngMode 
                  BackColor       =   &H00231C09&
                  Caption         =   "角度"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   84
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   780
               End
            End
            Begin VB.Label lblEntry 
               BackStyle       =   0  'Transparent
               Caption         =   "log 底数"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   83
               Top             =   1815
               Width           =   735
            End
            Begin VB.Label lblEntry 
               BackStyle       =   0  'Transparent
               Caption         =   "保留到小数点后"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   75
               Top             =   1350
               Width           =   1335
            End
            Begin VB.Label lblEntry 
               BackStyle       =   0  'Transparent
               Caption         =   "位小数"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   2040
               TabIndex        =   74
               Top             =   1350
               Width           =   615
            End
         End
         Begin VB.TextBox txtAnswer 
            BackColor       =   &H00231C09&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   2325
            Left            =   675
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   65
            Top             =   2415
            Width           =   3705
         End
         Begin VB.Label LBITEM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "计算器"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   1
            Left            =   2280
            TabIndex        =   77
            Top             =   150
            Width           =   540
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输入完整表达式程式即可"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   720
            TabIndex        =   76
            Top             =   1200
            Width           =   1980
         End
      End
      Begin VB.PictureBox PF 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   9375
         Index           =   11
         Left            =   0
         ScaleHeight     =   625
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   345
         TabIndex        =   280
         Top             =   7800
         Visible         =   0   'False
         Width           =   5175
         Begin ICEE.ICEE_WIN8 IWG 
            Height          =   2055
            Index           =   0
            Left            =   120
            TabIndex        =   282
            Top             =   1080
            Width           =   4740
            _ExtentX        =   8996
            _ExtentY        =   3863
         End
         Begin ICEE.ICEE_WIN8 IWG 
            Height          =   2055
            Index           =   1
            Left            =   120
            TabIndex        =   283
            Top             =   3240
            Width           =   4740
            _ExtentX        =   8996
            _ExtentY        =   3863
         End
         Begin ICEE.ICEE_WIN8 IWG 
            Height          =   2055
            Index           =   2
            Left            =   120
            TabIndex        =   284
            Top             =   5400
            Width           =   4740
            _ExtentX        =   8996
            _ExtentY        =   3863
         End
         Begin ICEE.ICEE_WIN8 IWG 
            Height          =   1500
            Index           =   3
            Left            =   120
            TabIndex        =   286
            Top             =   7560
            Width           =   1500
            _ExtentX        =   8996
            _ExtentY        =   3201
         End
         Begin ICEE.ICEE_WIN8 IWG 
            Height          =   1500
            Index           =   4
            Left            =   1725
            TabIndex        =   287
            Top             =   7560
            Width           =   1500
            _ExtentX        =   3493
            _ExtentY        =   3201
         End
         Begin VB.PictureBox PBK 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00586E74&
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   120
            ScaleHeight     =   65
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   65
            TabIndex        =   285
            Top             =   120
            Width           =   975
         End
         Begin ICEE.ICEE_WIN8 IWG 
            Height          =   1500
            Index           =   5
            Left            =   3360
            TabIndex        =   288
            Top             =   7560
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   2646
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ICEE新鲜事"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   33
            Left            =   1320
            TabIndex        =   281
            Top             =   360
            Width           =   1620
         End
      End
      Begin VB.PictureBox PicZoom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H005B1FC0&
         BorderStyle     =   0  'None
         Height          =   9375
         Left            =   3240
         ScaleHeight     =   625
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   340
         TabIndex        =   78
         Top             =   7560
         Visible         =   0   'False
         Width           =   5100
         Begin VB.Timer tmrZoom 
            Enabled         =   0   'False
            Interval        =   50
            Left            =   120
            Top             =   120
         End
         Begin VB.Label LA 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   4
            Left            =   4560
            TabIndex        =   270
            Top             =   9000
            Width           =   405
         End
         Begin VB.Image IMCZ 
            Height          =   900
            Index           =   1
            Left            =   240
            Top             =   7920
            Width           =   900
         End
         Begin VB.Image IMCZ 
            Height          =   900
            Index           =   0
            Left            =   240
            Top             =   6840
            Width           =   900
         End
         Begin VB.Image IWILLBK 
            Height          =   855
            Left            =   240
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.PictureBox PICCPU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         FillColor       =   &H0005FFFE&
         ForeColor       =   &H0005FFFE&
         Height          =   4875
         Left            =   3120
         ScaleHeight     =   325
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   321
         TabIndex        =   212
         Top             =   9000
         Visible         =   0   'False
         Width           =   4815
         Begin VB.PictureBox PCPU 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H00AD7900&
            ForeColor       =   &H00AD7900&
            Height          =   1455
            Left            =   600
            ScaleHeight     =   1455
            ScaleWidth      =   3255
            TabIndex        =   271
            Top             =   3000
            Width           =   3255
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "无法获取"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   32
            Left            =   2040
            TabIndex        =   275
            Top             =   1920
            Width           =   720
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "无法获取"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   23
            Left            =   2040
            TabIndex        =   274
            Top             =   1680
            Width           =   720
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ICEE已使用页面文件:"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   273
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ICEE已使用内存:"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   272
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CPU温度"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   230
            Top             =   840
            Width           =   735
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CPU使用"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   222
            Top             =   240
            Width           =   735
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "交换区"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   14
            Left            =   1320
            TabIndex        =   221
            Top             =   240
            Width           =   540
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "内存使用"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   5
            Left            =   1320
            TabIndex        =   220
            Top             =   840
            Width           =   720
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100%"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Index           =   9
            Left            =   360
            TabIndex        =   219
            Top             =   480
            Width           =   810
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "75%"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Index           =   12
            Left            =   1800
            TabIndex        =   218
            Top             =   480
            Width           =   630
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "15°C"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Index           =   41
            Left            =   360
            TabIndex        =   217
            Top             =   1080
            Width           =   675
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "750MB / 1024MB"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   42
            Left            =   2880
            TabIndex        =   216
            Top             =   360
            Width           =   1545
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "750MB / 1024MB"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   43
            Left            =   2880
            TabIndex        =   215
            Top             =   600
            Width           =   1545
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "70%"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Index           =   45
            Left            =   1800
            TabIndex        =   214
            Top             =   1080
            Width           =   630
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1530MB / 7777 MB"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   46
            Left            =   2880
            TabIndex        =   213
            Top             =   1200
            Width           =   1710
         End
      End
      Begin VB.PictureBox PF 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00231C09&
         BorderStyle     =   0  'None
         Height          =   1575
         Index           =   2
         Left            =   0
         ScaleHeight     =   105
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   340
         TabIndex        =   95
         Top             =   2640
         Visible         =   0   'False
         Width           =   5100
         Begin ICEE.ICEE_KEY ICL 
            Height          =   495
            Index           =   2
            Left            =   720
            TabIndex        =   96
            Top             =   840
            Width           =   1695
            _ExtentX        =   3625
            _ExtentY        =   661
         End
         Begin ICEE.ICEE_KEY ICL 
            Height          =   495
            Index           =   3
            Left            =   2640
            TabIndex        =   97
            Top             =   840
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            X1              =   0
            X2              =   344
            Y1              =   48
            Y2              =   48
         End
         Begin VB.Label LA 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "天啊,你的IE已经被甩出几条街了,想体验更精彩的ICEE吗,马上将IE升级到IE8吧,更精彩的内容等着你!"
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   15
            Left            =   240
            TabIndex        =   98
            Top             =   240
            Width           =   4410
            WordWrap        =   -1  'True
         End
         Begin VB.Image IMEND 
            Height          =   255
            Left            =   4800
            ToolTipText     =   "关闭"
            Top             =   60
            Width           =   255
         End
      End
      Begin VB.PictureBox PF 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00AD7900&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   8
         Left            =   2640
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   161
         TabIndex        =   276
         Top             =   7680
         Visible         =   0   'False
         Width           =   2415
         Begin VB.Image K 
            Height          =   240
            Index           =   4
            Left            =   1200
            Picture         =   "主窗体.frx":4ECD
            Top             =   120
            Width           =   240
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "118  KB"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   13
            Left            =   480
            TabIndex        =   278
            Top             =   120
            Width           =   675
         End
         Begin VB.Image K 
            Height          =   240
            Index           =   7
            Left            =   120
            Picture         =   "主窗体.frx":5257
            Top             =   120
            Width           =   240
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "500 KB"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   44
            Left            =   1560
            TabIndex        =   277
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.PictureBox Picd 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0C0C0&
         Height          =   5940
         Left            =   2040
         ScaleHeight     =   396
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   340
         TabIndex        =   14
         Top             =   9240
         Visible         =   0   'False
         Width           =   5100
         Begin ICEE.ICEE_KEY ICL 
            Height          =   450
            Index           =   5
            Left            =   3480
            TabIndex        =   159
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   794
         End
         Begin VB.PictureBox PICSLIDE 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00514E50&
            BorderStyle     =   0  'None
            Height          =   4620
            Left            =   0
            ScaleHeight     =   308
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   153
            TabIndex        =   155
            Top             =   1080
            Width           =   2295
            Begin VB.FileListBox FILHIDDEN 
               Height          =   3330
               Left            =   3000
               TabIndex        =   157
               Top             =   0
               Visible         =   0   'False
               Width           =   1695
            End
            Begin ICEE.ICEE_WIN8 IPLAY 
               Height          =   4695
               Index           =   0
               Left            =   0
               TabIndex        =   156
               Top             =   0
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   8281
            End
         End
         Begin ICEE.IHScroll HScroll1 
            Height          =   225
            Left            =   0
            TabIndex        =   162
            Top             =   5715
            Width           =   5100
            _ExtentX        =   8705
            _ExtentY        =   397
            MinV            =   0
            MaxV            =   20
            Value           =   0
            SmallChange     =   1
            LargeChange     =   10
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   27
            Left            =   3000
            TabIndex        =   163
            Top             =   3000
            Width           =   780
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "APP.PATH"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   24
            Left            =   120
            TabIndex        =   158
            Top             =   120
            Width           =   720
         End
         Begin VB.Image IMR 
            Height          =   510
            Left            =   120
            Top             =   360
            Width           =   1275
         End
      End
      Begin VB.PictureBox PICSER 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H001F1F1F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   0
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   340
         TabIndex        =   50
         Top             =   1800
         Visible         =   0   'False
         Width           =   5100
         Begin ICEE.IList LISTBAIDU 
            Height          =   2460
            Left            =   15
            TabIndex        =   126
            Top             =   375
            Visible         =   0   'False
            Width           =   5070
            _ExtentX        =   8943
            _ExtentY        =   4339
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ItemHeight      =   18
         End
         Begin VB.TextBox TXTBAIDU 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   240
            TabIndex        =   51
            Text            =   "<请输入关键词>"
            Top             =   120
            Width           =   4215
         End
         Begin VB.Image IMBAIDU 
            Height          =   255
            Left            =   4680
            ToolTipText     =   "关闭搜索框"
            Top             =   120
            Width           =   255
         End
      End
      Begin MSWinsockLib.Winsock WinC 
         Left            =   4320
         Top             =   4440
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.PictureBox PDB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00B6B067&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6000
         Left            =   4320
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   400
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   340
         TabIndex        =   18
         Top             =   8880
         Width           =   5100
         Begin ICEE.ICEE_KEY ICL 
            Height          =   570
            Index           =   6
            Left            =   1560
            TabIndex        =   164
            Top             =   3960
            Width           =   1215
            _ExtentX        =   1931
            _ExtentY        =   794
         End
         Begin VB.Label LA 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00383636&
            Height          =   345
            Index           =   2
            Left            =   1275
            TabIndex        =   80
            Top             =   4050
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label 提示信息 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "正在初始化内存"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   3600
            TabIndex        =   19
            Top             =   360
            UseMnemonic     =   0   'False
            Width           =   1260
         End
      End
      Begin VB.PictureBox PICTIME 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H006C7509&
         BorderStyle     =   0  'None
         Height          =   1065
         Left            =   3420
         ScaleHeight     =   71
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   100
         Top             =   0
         Width           =   1695
         Begin VB.Image UNME 
            Height          =   240
            Left            =   1320
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗体.frx":55E1
            Top             =   120
            Width           =   240
         End
         Begin VB.Image MINIME 
            Height          =   240
            Left            =   960
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗体.frx":596B
            ToolTipText     =   "最小化"
            Top             =   120
            Width           =   240
         End
         Begin VB.Image BACKME 
            Height          =   240
            Left            =   240
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗体.frx":5CF5
            ToolTipText     =   "返回主面板"
            Top             =   120
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image SETME 
            Height          =   240
            Left            =   600
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗体.frx":607F
            ToolTipText     =   "快速设置"
            Top             =   120
            Width           =   240
         End
      End
      Begin VB.PictureBox PICAD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00231C09&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   9360
         Left            =   3840
         ScaleHeight     =   624
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   340
         TabIndex        =   42
         Top             =   8520
         Width           =   5100
         Begin VB.Timer TMAD 
            Enabled         =   0   'False
            Interval        =   50
            Left            =   120
            Top             =   120
         End
      End
      Begin VB.PictureBox PR 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0028D985&
         BorderStyle     =   0  'None
         Height          =   75
         Left            =   0
         ScaleHeight     =   5
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   81
         Top             =   7560
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox PICTOOL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00201400&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   340
         TabIndex        =   24
         Top             =   8160
         Width           =   5100
         Begin VB.PictureBox PF 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00201400&
            BorderStyle     =   0  'None
            FillStyle       =   3  'Vertical Line
            Height          =   375
            Index           =   6
            Left            =   4440
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   25
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   340
            TabIndex        =   25
            Top             =   360
            Visible         =   0   'False
            Width           =   5100
            Begin VB.Image ld2 
               Height          =   360
               Left            =   0
               Picture         =   "主窗体.frx":6409
               ToolTipText     =   "放大"
               Top             =   0
               Visible         =   0   'False
               Width           =   360
            End
            Begin VB.Image pe1 
               Height          =   360
               Left            =   4680
               Picture         =   "主窗体.frx":6B73
               ToolTipText     =   "自定义图片(你可以定义7张)"
               Top             =   0
               Width           =   360
            End
            Begin VB.Image ld1 
               Height          =   360
               Left            =   0
               Picture         =   "主窗体.frx":72DD
               ToolTipText     =   "放大"
               Top             =   0
               Width           =   360
            End
            Begin VB.Image ld3 
               Height          =   360
               Left            =   0
               Picture         =   "主窗体.frx":7A47
               ToolTipText     =   "放大"
               Top             =   0
               Visible         =   0   'False
               Width           =   360
            End
            Begin VB.Image pe3 
               Height          =   360
               Left            =   4680
               Picture         =   "主窗体.frx":81B1
               ToolTipText     =   "自定义图片(你可以定义7张)"
               Top             =   0
               Visible         =   0   'False
               Width           =   360
            End
            Begin VB.Image pe2 
               Height          =   360
               Left            =   4680
               Picture         =   "主窗体.frx":891B
               ToolTipText     =   "自定义图片(你可以定义7张)"
               Top             =   0
               Visible         =   0   'False
               Width           =   360
            End
         End
         Begin VB.Image IMCLIP 
            Height          =   240
            Left            =   840
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗体.frx":9085
            Top             =   60
            Width           =   240
         End
         Begin VB.Image IMSIN 
            Height          =   240
            Left            =   4680
            Picture         =   "主窗体.frx":940F
            Top             =   75
            Width           =   240
         End
         Begin VB.Image IMGUSB 
            Height          =   240
            Left            =   480
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗体.frx":9799
            Top             =   45
            Width           =   240
         End
         Begin VB.Image IMCPU 
            Height          =   240
            Left            =   120
            Picture         =   "主窗体.frx":9B23
            Top             =   75
            Width           =   240
         End
      End
      Begin VB.PictureBox IMJ 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H004E4E4E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4680
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   45
         ToolTipText     =   "关闭"
         Top             =   2880
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox PLOGU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00201400&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3105
         Left            =   1680
         ScaleHeight     =   207
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   207
         TabIndex        =   32
         Top             =   9600
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.PictureBox PLOGO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00201400&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3105
         Left            =   1680
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   207
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   207
         TabIndex        =   29
         Top             =   9600
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   1920
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   89
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   89
         TabIndex        =   31
         Top             =   10560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   3360
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   89
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   89
         TabIndex        =   30
         Top             =   10560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Timer TMO 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   10920
         Top             =   615
      End
      Begin VB.Timer Timefriend 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1800
         Top             =   960
      End
      Begin VB.Timer Timetool 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   10905
         Top             =   120
      End
      Begin VB.Timer TMRZ 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   10905
         Top             =   2520
      End
      Begin VB.Timer TmrBK 
         Interval        =   30
         Left            =   10920
         Top             =   2040
      End
      Begin VB.Timer TMP 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   10905
         Top             =   1095
      End
      Begin VB.Timer Timers 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   2280
         Top             =   960
      End
      Begin VB.PictureBox pc 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H003F3B36&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   600
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   300
         TabIndex        =   2
         Top             =   8520
         Width           =   4500
         Begin VB.Label lbthing 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "欢迎使用1.22全新版本,更多精彩等你发现"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   0
            OLEDropMode     =   1  'Manual
            TabIndex        =   3
            Top             =   120
            Width           =   4290
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox PF 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00201400&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5865
         Index           =   3
         Left            =   0
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   391
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   340
         TabIndex        =   11
         Top             =   3000
         Width           =   5100
         Begin VB.PictureBox PicUse 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   5505
            Left            =   360
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   367
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   340
            TabIndex        =   101
            Top             =   720
            Width           =   5100
            Begin VB.PictureBox PF 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H007E5502&
               BorderStyle     =   0  'None
               Height          =   735
               Index           =   17
               Left            =   0
               ScaleHeight     =   49
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   340
               TabIndex        =   308
               Top             =   0
               Visible         =   0   'False
               Width           =   5100
               Begin ICEE.ICEE_KEY ICOCO 
                  Height          =   495
                  Index           =   0
                  Left            =   120
                  TabIndex        =   309
                  Top             =   120
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   873
               End
               Begin ICEE.ICEE_KEY ICOCO 
                  Height          =   495
                  Index           =   1
                  Left            =   1440
                  TabIndex        =   310
                  Top             =   120
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   873
               End
               Begin ICEE.ICEE_KEY ICOCO 
                  Height          =   495
                  Index           =   2
                  Left            =   2880
                  TabIndex        =   316
                  Top             =   120
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   873
               End
               Begin VB.Image ICLOS 
                  Height          =   240
                  Left            =   4680
                  OLEDropMode     =   1  'Manual
                  Picture         =   "主窗体.frx":9EAD
                  Top             =   240
                  Width           =   240
               End
            End
            Begin VB.PictureBox PICCLIP 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00231C09&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   5520
               Left            =   2640
               OLEDropMode     =   1  'Manual
               ScaleHeight     =   368
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   340
               TabIndex        =   102
               Top             =   4320
               Visible         =   0   'False
               Width           =   5100
               Begin VB.PictureBox ImgPreview 
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   3015
                  Left            =   1200
                  ScaleHeight     =   201
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   265
                  TabIndex        =   104
                  Top             =   4080
                  Visible         =   0   'False
                  Width           =   3975
               End
               Begin MSComctlLib.ListView ListView1 
                  Height          =   2085
                  Left            =   195
                  TabIndex        =   105
                  Top             =   3285
                  Width           =   4635
                  _ExtentX        =   8176
                  _ExtentY        =   3678
                  View            =   3
                  Arrange         =   1
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  _Version        =   393217
                  Icons           =   "iCLIP"
                  SmallIcons      =   "iCLIP"
                  ForeColor       =   0
                  BackColor       =   16777215
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   0
               End
               Begin MSComctlLib.ImageList iCLIP 
                  Left            =   3480
                  Top             =   0
                  _ExtentX        =   1005
                  _ExtentY        =   1005
                  BackColor       =   -2147483643
                  ImageWidth      =   32
                  ImageHeight     =   32
                  MaskColor       =   12632256
                  _Version        =   393216
                  BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
                     NumListImages   =   2
                     BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "主窗体.frx":A237
                        Key             =   ""
                     EndProperty
                     BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "主窗体.frx":AF11
                        Key             =   ""
                     EndProperty
                  EndProperty
               End
               Begin VB.PictureBox PVP 
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H00231C09&
                  BorderStyle     =   0  'None
                  Height          =   2175
                  Left            =   120
                  ScaleHeight     =   145
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   316
                  TabIndex        =   107
                  Top             =   720
                  Width           =   4740
                  Begin ICEE.ICEE_COMMAND ICC 
                     Height          =   540
                     Index           =   0
                     Left            =   0
                     TabIndex        =   108
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   1890
                     _ExtentX        =   3334
                     _ExtentY        =   873
                  End
                  Begin ICEE.ICEE_COMMAND ICC 
                     Height          =   540
                     Index           =   1
                     Left            =   0
                     TabIndex        =   109
                     Top             =   480
                     Visible         =   0   'False
                     Width           =   1890
                     _ExtentX        =   3201
                     _ExtentY        =   873
                  End
                  Begin ICEE.ICEE_COMMAND ICC 
                     Height          =   600
                     Index           =   2
                     Left            =   0
                     TabIndex        =   110
                     Top             =   960
                     Visible         =   0   'False
                     Width           =   1890
                     _ExtentX        =   3201
                     _ExtentY        =   873
                  End
                  Begin ICEE.ICEE_COMMAND ICC 
                     Height          =   675
                     Index           =   3
                     Left            =   0
                     TabIndex        =   111
                     Top             =   1500
                     Visible         =   0   'False
                     Width           =   1890
                     _ExtentX        =   3201
                     _ExtentY        =   873
                  End
                  Begin VB.Image IPR 
                     Appearance      =   0  'Flat
                     Enabled         =   0   'False
                     Height          =   2115
                     Left            =   1560
                     Stretch         =   -1  'True
                     Top             =   120
                     Width           =   1815
                  End
               End
               Begin ICEE.ICEE_COMMAND ICT 
                  Height          =   435
                  Left            =   120
                  TabIndex        =   103
                  Top             =   2490
                  Visible         =   0   'False
                  Width           =   4755
                  _ExtentX        =   3201
                  _ExtentY        =   873
               End
               Begin VB.TextBox txtText 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   1785
                  Left            =   120
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  OLEDragMode     =   1  'Automatic
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   106
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   4740
               End
               Begin VB.Image IMCLP 
                  Height          =   510
                  Left            =   120
                  Top             =   0
                  Width           =   1275
               End
               Begin VB.Label LBITEM 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "剪切板"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   180
                  Index           =   12
                  Left            =   2280
                  TabIndex        =   112
                  Top             =   165
                  Width           =   540
               End
            End
            Begin VB.PictureBox PF 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   11175
               Index           =   4
               Left            =   0
               ScaleHeight     =   745
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   345
               TabIndex        =   128
               Top             =   1320
               Visible         =   0   'False
               Width           =   5175
               Begin ICEE.ICEE_WIN8 IW 
                  Height          =   1620
                  Index           =   10
                  Left            =   3420
                  TabIndex        =   138
                  Top             =   7800
                  Width           =   1620
                  _ExtentX        =   2858
                  _ExtentY        =   2858
               End
               Begin ICEE.ICEE_WIN8 IW 
                  Height          =   3315
                  Index           =   0
                  Left            =   30
                  TabIndex        =   140
                  Top             =   45
                  Width           =   3315
                  _ExtentX        =   5847
                  _ExtentY        =   5847
                  Begin ICEE.ucScrollbar SHRO 
                     Height          =   1695
                     Left            =   4680
                     TabIndex        =   141
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   255
                     _ExtentX        =   450
                     _ExtentY        =   2990
                  End
                  Begin VB.Image IMKK 
                     Height          =   1215
                     Left            =   1080
                     Top             =   960
                     Width           =   1215
                  End
               End
               Begin ICEE.ICEE_WIN8 IW 
                  Height          =   1620
                  Index           =   7
                  Left            =   3420
                  TabIndex        =   129
                  Top             =   3450
                  Width           =   1620
                  _ExtentX        =   2858
                  _ExtentY        =   2858
               End
               Begin ICEE.ICEE_WIN8 IW 
                  Height          =   1620
                  Index           =   1
                  Left            =   1725
                  TabIndex        =   130
                  Top             =   3450
                  Width           =   1620
                  _ExtentX        =   2858
                  _ExtentY        =   2858
               End
               Begin ICEE.ICEE_WIN8 IW 
                  Height          =   1620
                  Index           =   2
                  Left            =   1725
                  TabIndex        =   131
                  Top             =   7800
                  Width           =   1620
                  _ExtentX        =   2858
                  _ExtentY        =   2858
               End
               Begin ICEE.ICEE_WIN8 IW 
                  Height          =   1620
                  Index           =   3
                  Left            =   30
                  TabIndex        =   132
                  Top             =   3450
                  Width           =   1620
                  _ExtentX        =   2858
                  _ExtentY        =   2858
               End
               Begin ICEE.ICEE_WIN8 IW 
                  Height          =   1620
                  Index           =   4
                  Left            =   3420
                  TabIndex        =   133
                  Top             =   1740
                  Width           =   1620
                  _ExtentX        =   2858
                  _ExtentY        =   2858
               End
               Begin ICEE.ICEE_WIN8 IW 
                  Height          =   1620
                  Index           =   5
                  Left            =   30
                  TabIndex        =   134
                  Top             =   7800
                  Width           =   1620
                  _ExtentX        =   2858
                  _ExtentY        =   2858
               End
               Begin ICEE.ICEE_WIN8 IW 
                  Height          =   1620
                  Index           =   6
                  Left            =   3420
                  TabIndex        =   135
                  Top             =   9495
                  Width           =   1620
                  _ExtentX        =   2858
                  _ExtentY        =   2858
               End
               Begin ICEE.ICEE_WIN8 IW 
                  Height          =   1620
                  Index           =   8
                  Left            =   3420
                  TabIndex        =   136
                  Top             =   45
                  Width           =   1620
                  _ExtentX        =   2858
                  _ExtentY        =   2858
                  Begin VB.Label LA 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "周日"
                     BeginProperty Font 
                        Name            =   "微软雅黑"
                        Size            =   9
                        Charset         =   134
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   255
                     Index           =   17
                     Left            =   120
                     TabIndex        =   311
                     Top             =   1320
                     Width           =   360
                  End
                  Begin VB.Label LA 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "蛇年"
                     BeginProperty Font 
                        Name            =   "微软雅黑"
                        Size            =   9
                        Charset         =   134
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   255
                     Index           =   26
                     Left            =   1200
                     TabIndex        =   161
                     Top             =   1320
                     Width           =   360
                  End
                  Begin VB.Label LA 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "2013-12-1"
                     BeginProperty Font 
                        Name            =   "微软雅黑"
                        Size            =   9
                        Charset         =   134
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   255
                     Index           =   25
                     Left            =   120
                     TabIndex        =   160
                     Top             =   120
                     Width           =   885
                  End
               End
               Begin ICEE.ICEE_WIN8 IW 
                  Height          =   1620
                  Index           =   9
                  Left            =   30
                  TabIndex        =   137
                  Top             =   9495
                  Width           =   3315
                  _ExtentX        =   5847
                  _ExtentY        =   2858
               End
               Begin ICEE.ICEE_WIN8 IW 
                  Height          =   2550
                  Index           =   11
                  Left            =   30
                  TabIndex        =   139
                  Top             =   5175
                  Width           =   5010
                  _ExtentX        =   8837
                  _ExtentY        =   4498
               End
            End
            Begin VB.Label LBITEM 
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00C0C0C0&
               Height          =   1350
               Index           =   11
               Left            =   3720
               TabIndex        =   122
               Top             =   720
               Width           =   1350
            End
            Begin VB.Label LBITEM 
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00C0C0C0&
               Height          =   1350
               Index           =   10
               Left            =   2040
               TabIndex        =   121
               Top             =   720
               Width           =   1350
            End
            Begin VB.Label LBITEM 
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00C0C0C0&
               Height          =   1350
               Index           =   7
               Left            =   360
               TabIndex        =   120
               Top             =   600
               Width           =   1350
            End
            Begin VB.Label LBITEM 
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00C0C0C0&
               Height          =   1350
               Index           =   0
               Left            =   360
               TabIndex        =   119
               Top             =   2160
               Width           =   1350
            End
            Begin VB.Label LBITEM 
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00C0C0C0&
               Height          =   1350
               Index           =   8
               Left            =   3720
               TabIndex        =   118
               Top             =   2160
               Width           =   1350
            End
            Begin VB.Label LBITEM 
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00C0C0C0&
               Height          =   1350
               Index           =   6
               Left            =   2040
               TabIndex        =   117
               Top             =   3720
               Width           =   1350
            End
            Begin VB.Label LBITEM 
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00C0C0C0&
               Height          =   1350
               Index           =   5
               Left            =   2040
               TabIndex        =   116
               Top             =   2160
               Width           =   1350
            End
            Begin VB.Label LBITEM 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "主菜单"
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   3
               Left            =   2280
               TabIndex        =   115
               Top             =   165
               Width           =   540
            End
            Begin VB.Label LBITEM 
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00C0C0C0&
               Height          =   1350
               Index           =   4
               Left            =   360
               TabIndex        =   114
               Top             =   3720
               Width           =   1350
            End
            Begin VB.Label LA 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "微软雅黑"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   3
               Left            =   1440
               TabIndex        =   113
               Top             =   480
               Width           =   135
            End
         End
         Begin VB.PictureBox PICIM 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H007A7417&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   6570
            Left            =   120
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   438
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   340
            TabIndex        =   12
            Top             =   720
            Width           =   5100
            Begin VB.PictureBox PICIG 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00633F0E&
               BorderStyle     =   0  'None
               FillColor       =   &H00383537&
               FillStyle       =   2  'Horizontal Line
               ForeColor       =   &H00FFFFFF&
               Height          =   5550
               Left            =   -2040
               OLEDropMode     =   1  'Manual
               ScaleHeight     =   370
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   340
               TabIndex        =   27
               Top             =   6240
               Visible         =   0   'False
               Width           =   5100
               Begin ICEE.ICEE_KEY ICL 
                  Height          =   495
                  Index           =   13
                  Left            =   840
                  TabIndex        =   189
                  Top             =   4440
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   873
               End
               Begin VB.ListBox LSTBOX 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00000000&
                  Height          =   2220
                  ItemData        =   "主窗体.frx":BBEB
                  Left            =   840
                  List            =   "主窗体.frx":BBED
                  TabIndex        =   61
                  Top             =   1920
                  Width           =   3285
               End
               Begin VB.ComboBox TXTBOX 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00000000&
                  Height          =   300
                  Left            =   840
                  TabIndex        =   28
                  Top             =   1200
                  Width           =   3255
               End
               Begin ICEE.ICEE_KEY ICL 
                  Height          =   495
                  Index           =   14
                  Left            =   3000
                  TabIndex        =   190
                  Top             =   4440
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   873
               End
            End
            Begin VB.PictureBox PICBUG 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00828637&
               BorderStyle     =   0  'None
               ForeColor       =   &H0099FFFF&
               Height          =   5550
               Left            =   -3000
               OLEDropMode     =   1  'Manual
               ScaleHeight     =   370
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   340
               TabIndex        =   22
               Top             =   6240
               Visible         =   0   'False
               Width           =   5100
               Begin ICEE.ICEE_KEY ICL 
                  Height          =   495
                  Index           =   11
                  Left            =   720
                  TabIndex        =   187
                  Top             =   4200
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   873
               End
               Begin VB.TextBox TXTBUG 
                  Appearance      =   0  'Flat
                  BackColor       =   &H007A7417&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H00FFFFFF&
                  Height          =   3285
                  Left            =   720
                  MaxLength       =   500
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   23
                  Top             =   720
                  Width           =   3735
               End
               Begin ICEE.ICEE_KEY ICL 
                  Height          =   495
                  Index           =   12
                  Left            =   3000
                  TabIndex        =   188
                  Top             =   4200
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   873
               End
            End
            Begin VB.PictureBox PICLO 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00383537&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   6015
               Left            =   120
               OLEDropMode     =   1  'Manual
               ScaleHeight     =   401
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   340
               TabIndex        =   46
               Top             =   0
               Width           =   5100
               Begin VB.PictureBox PICNET 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H007A7417&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   5100
                  Left            =   0
                  OLEDropMode     =   1  'Manual
                  ScaleHeight     =   340
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   340
                  TabIndex        =   52
                  Top             =   5400
                  Visible         =   0   'False
                  Width           =   5100
                  Begin VB.TextBox Text3 
                     BackColor       =   &H00373637&
                     BorderStyle     =   0  'None
                     BeginProperty DataFormat 
                        Type            =   1
                        Format          =   "0"
                        HaveTrueFalseNull=   0
                        FirstDayOfWeek  =   0
                        FirstWeekOfYear =   0
                        LCID            =   2052
                        SubFormatType   =   1
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   255
                     IMEMode         =   3  'DISABLE
                     Left            =   360
                     MaxLength       =   15
                     TabIndex        =   79
                     Text            =   "<输入IP>"
                     Top             =   4500
                     Width           =   2715
                  End
                  Begin VB.PictureBox PF 
                     AutoRedraw      =   -1  'True
                     AutoSize        =   -1  'True
                     BackColor       =   &H00CCF4F3&
                     BorderStyle     =   0  'None
                     Height          =   1560
                     Index           =   10
                     Left            =   360
                     ScaleHeight     =   104
                     ScaleMode       =   3  'Pixel
                     ScaleWidth      =   188
                     TabIndex        =   57
                     Top             =   2040
                     Width           =   2820
                     Begin VB.ListBox lstRes 
                        Appearance      =   0  'Flat
                        BackColor       =   &H00373637&
                        ForeColor       =   &H00FFFFFF&
                        Height          =   1590
                        IntegralHeight  =   0   'False
                        ItemData        =   "主窗体.frx":BBEF
                        Left            =   -15
                        List            =   "主窗体.frx":BBF1
                        TabIndex        =   58
                        Top             =   -15
                        Width           =   2850
                     End
                  End
                  Begin VB.TextBox txtFrom 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00373637&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H00FFFFFF&
                     Height          =   195
                     Left            =   2445
                     TabIndex        =   54
                     Text            =   "0"
                     Top             =   720
                     Width           =   495
                  End
                  Begin VB.TextBox txtTo 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00373637&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H00FFFFFF&
                     Height          =   285
                     Left            =   2460
                     TabIndex        =   53
                     Text            =   "99"
                     Top             =   960
                     Width           =   495
                  End
                  Begin ICEE.ICEE_KEY ICL 
                     Height          =   570
                     Index           =   7
                     Left            =   1080
                     TabIndex        =   165
                     Top             =   1320
                     Width           =   1215
                     _extentx        =   1931
                     _extenty        =   794
                  End
                  Begin VB.Label LA 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "输入服务器地址"
                     ForeColor       =   &H00FFFFFF&
                     Height          =   180
                     Index           =   28
                     Left            =   360
                     TabIndex        =   167
                     Top             =   4200
                     Width           =   1260
                  End
                  Begin VB.Label lblTo 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "结束 000.000.000."
                     ForeColor       =   &H00FFFFFF&
                     Height          =   195
                     Left            =   720
                     TabIndex        =   56
                     Top             =   960
                     Width           =   1575
                  End
                  Begin VB.Label lblFrom 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "起始 000.000.000."
                     ForeColor       =   &H00FFFFFF&
                     Height          =   195
                     Left            =   720
                     TabIndex        =   55
                     Top             =   720
                     Width           =   1575
                  End
               End
               Begin VB.TextBox Text2 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   1
                  EndProperty
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   10.5
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  HideSelection   =   0   'False
                  IMEMode         =   3  'DISABLE
                  Left            =   1080
                  MaxLength       =   16
                  TabIndex        =   48
                  Top             =   3240
                  Width           =   2955
               End
               Begin VB.TextBox Text1 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   1
                  EndProperty
                  BeginProperty Font 
                     Name            =   "微软雅黑"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  IMEMode         =   3  'DISABLE
                  Left            =   1080
                  MaxLength       =   12
                  TabIndex        =   47
                  Text            =   "<请输入ID>"
                  Top             =   2370
                  Width           =   2955
               End
               Begin MSWinsockLib.Winsock Winsock1 
                  Left            =   1800
                  Top             =   360
                  _ExtentX        =   741
                  _ExtentY        =   741
                  _Version        =   393216
                  RemoteHost      =   "127.0.0.1"
                  RemotePort      =   6000
               End
               Begin ICEE.ICEE_KEY ICL 
                  Height          =   570
                  Index           =   1
                  Left            =   3000
                  TabIndex        =   166
                  Top             =   4080
                  Width           =   1215
                  _extentx        =   1931
                  _extenty        =   794
               End
               Begin ICEE.ICHECK ICK 
                  Height          =   375
                  Index           =   2
                  Left            =   840
                  TabIndex        =   170
                  Top             =   4500
                  Width           =   1935
                  _extentx        =   3413
                  _extenty        =   661
               End
               Begin ICEE.ICHECK ICK 
                  Height          =   375
                  Index           =   1
                  Left            =   840
                  TabIndex        =   169
                  Top             =   4140
                  Width           =   1935
                  _extentx        =   3413
                  _extenty        =   873
               End
               Begin ICEE.ICHECK ICK 
                  Height          =   375
                  Index           =   0
                  Left            =   840
                  TabIndex        =   168
                  Top             =   3780
                  Width           =   1935
                  _extentx        =   3413
                  _extenty        =   661
               End
               Begin VB.Image IMG_NT 
                  Height          =   630
                  Left            =   120
                  Top             =   120
                  Width           =   630
               End
               Begin VB.Label LBITEM 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "请登陆"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   180
                  Index           =   2
                  Left            =   2280
                  TabIndex        =   49
                  Top             =   165
                  Width           =   540
               End
            End
            Begin VB.Timer BuddyUpdater 
               Enabled         =   0   'False
               Interval        =   15000
               Left            =   2520
               Top             =   6000
            End
            Begin VB.PictureBox PICFI 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H007A7417&
               BorderStyle     =   0  'None
               ForeColor       =   &H00E0E0E0&
               Height          =   5550
               Left            =   -600
               OLEDropMode     =   1  'Manual
               ScaleHeight     =   370
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   340
               TabIndex        =   26
               Top             =   6480
               Visible         =   0   'False
               Width           =   5100
               Begin MSComctlLib.ImageList USEZT 
                  Left            =   0
                  Top             =   0
                  _ExtentX        =   1005
                  _ExtentY        =   1005
                  BackColor       =   -2147483643
                  ImageWidth      =   48
                  ImageHeight     =   48
                  MaskColor       =   12632256
                  UseMaskColor    =   0   'False
                  _Version        =   393216
                  BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
                     NumListImages   =   5
                     BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "主窗体.frx":BBF3
                        Key             =   "ONLINE"
                     EndProperty
                     BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "主窗体.frx":D8CD
                        Key             =   "BUSY"
                     EndProperty
                     BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "主窗体.frx":F5A7
                        Key             =   "OFFLINE"
                     EndProperty
                     BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "主窗体.frx":11281
                        Key             =   "DNZ"
                     EndProperty
                     BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "主窗体.frx":12F5B
                        Key             =   "UNKNOW"
                     EndProperty
                  EndProperty
               End
               Begin VB.Label LBFN 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  ForeColor       =   &H00FFFFFF&
                  Height          =   225
                  Left            =   840
                  TabIndex        =   40
                  Top             =   3960
                  Width           =   3330
                  WordWrap        =   -1  'True
               End
               Begin VB.Label LBFW 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "未设置"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   180
                  Left            =   600
                  TabIndex        =   39
                  Top             =   3180
                  Width           =   540
               End
               Begin VB.Label LBFA 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "未设置"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   180
                  Left            =   2400
                  TabIndex        =   38
                  Top             =   2460
                  Width           =   540
               End
               Begin VB.Label LBFS 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "未设置"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   180
                  Left            =   600
                  TabIndex        =   37
                  Top             =   2460
                  Width           =   540
               End
               Begin VB.Label LBFE 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "未设置"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   180
                  Left            =   600
                  TabIndex        =   36
                  Top             =   1740
                  Width           =   540
               End
               Begin VB.Label LBFQ 
                  BackStyle       =   0  'Transparent
                  Caption         =   "未设置"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   180
                  Left            =   600
                  TabIndex        =   35
                  Top             =   1080
                  Width           =   3330
               End
            End
            Begin VB.PictureBox PICPASS 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H005034D3&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FFFFFF&
               Height          =   5550
               Left            =   -480
               OLEDropMode     =   1  'Manual
               ScaleHeight     =   370
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   340
               TabIndex        =   33
               Top             =   6120
               Visible         =   0   'False
               Width           =   5100
               Begin ICEE.ICEE_KEY ICL 
                  Height          =   495
                  Index           =   15
                  Left            =   1080
                  TabIndex        =   191
                  Top             =   3120
                  Width           =   1455
                  _extentx        =   2566
                  _extenty        =   873
               End
               Begin VB.TextBox TXTPASS 
                  Appearance      =   0  'Flat
                  BackColor       =   &H007A7417&
                  BorderStyle     =   0  'None
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   1
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   1080
                  MaxLength       =   16
                  TabIndex        =   34
                  Top             =   2880
                  Width           =   2895
               End
               Begin ICEE.ICEE_KEY ICL 
                  Height          =   495
                  Index           =   16
                  Left            =   2520
                  TabIndex        =   192
                  Top             =   3120
                  Width           =   1455
                  _extentx        =   2778
                  _extenty        =   873
               End
            End
            Begin ICEE.ICEE_KEY ICL 
               Height          =   615
               Index           =   8
               Left            =   3840
               TabIndex        =   193
               Top             =   600
               Width           =   1095
               _extentx        =   1931
               _extenty        =   1085
            End
            Begin VB.TextBox TXTSER 
               Appearance      =   0  'Flat
               BackColor       =   &H001B27C9&
               BorderStyle     =   0  'None
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   360
               MaxLength       =   12
               OLEDropMode     =   1  'Manual
               TabIndex        =   43
               Text            =   "请输入对方ID按回车键确认"
               Top             =   840
               Width           =   3375
            End
            Begin MSComctlLib.TreeView TreeView1 
               Height          =   3375
               Left            =   120
               TabIndex        =   13
               Top             =   1320
               Width           =   4815
               _ExtentX        =   8493
               _ExtentY        =   5953
               _Version        =   393217
               Indentation     =   265
               LabelEdit       =   1
               Sorted          =   -1  'True
               Style           =   1
               FullRowSelect   =   -1  'True
               SingleSel       =   -1  'True
               ImageList       =   "USEZT"
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               OLEDropMode     =   1
            End
            Begin VB.Shape SB 
               BackColor       =   &H001B27C9&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H001B27C9&
               BorderStyle     =   0  'Transparent
               Height          =   630
               Index           =   0
               Left            =   120
               Top             =   600
               Width           =   3735
            End
            Begin VB.Label LA 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "好友列表"
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   1
               Left            =   2280
               TabIndex        =   41
               Top             =   165
               Width           =   720
            End
            Begin VB.Label LBC 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "你还没有好友呢"
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Left            =   240
               TabIndex        =   21
               Top             =   4920
               Width           =   1260
            End
         End
         Begin VB.PictureBox pl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00201400&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   5625
            Left            =   0
            ScaleHeight     =   375
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   342
            TabIndex        =   15
            Top             =   360
            Width           =   5130
            Begin ICEE.ICEE_COMMAND ICZ 
               Height          =   450
               Index           =   3
               Left            =   120
               TabIndex        =   88
               Top             =   75
               Visible         =   0   'False
               Width           =   975
               _extentx        =   1720
               _extenty        =   873
            End
         End
      End
      Begin VB.PictureBox PICMU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H003F3B36&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   0
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   1
         ToolTipText     =   "主菜单"
         Top             =   8640
         Width           =   720
      End
      Begin VB.PictureBox PNZ 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H003F3B36&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   0
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   44
         ToolTipText     =   "主菜单"
         Top             =   8640
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label LA 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "New"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   16
         Left            =   4680
         TabIndex        =   292
         Top             =   1305
         Width           =   270
      End
      Begin VB.Image IMGEM 
         Height          =   240
         Left            =   3960
         OLEDropMode     =   1  'Manual
         Picture         =   "主窗体.frx":14C35
         ToolTipText     =   "未读消息"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Label lbzt 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00A5A5A5&
         Height          =   420
         Left            =   1350
         TabIndex        =   94
         Top             =   1200
         Width           =   330
      End
      Begin VB.Label LCO 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   3780
         TabIndex        =   93
         Top             =   1305
         Width           =   90
      End
      Begin VB.Image IMSKIN 
         Height          =   240
         Left            =   4320
         Picture         =   "主窗体.frx":14FBF
         Top             =   1275
         Width           =   240
      End
      Begin VB.Image IMAD 
         Height          =   360
         Left            =   4680
         OLEDropMode     =   1  'Manual
         ToolTipText     =   "查看海报"
         Top             =   2370
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image IMCHAT 
         Height          =   360
         Left            =   3600
         OLEDropMode     =   1  'Manual
         Top             =   2370
         Width           =   525
      End
      Begin VB.Image IMPIC 
         Height          =   360
         Left            =   2280
         OLEDropMode     =   1  'Manual
         Top             =   2370
         Width           =   525
      End
      Begin VB.Image IMMAIN 
         Height          =   360
         Left            =   960
         OLEDropMode     =   1  'Manual
         Top             =   2370
         Width           =   525
      End
      Begin VB.Image IMSER 
         Height          =   360
         Left            =   240
         OLEDropMode     =   1  'Manual
         ToolTipText     =   "搜索"
         Top             =   2370
         Width           =   345
      End
      Begin VB.Image E2 
         Enabled         =   0   'False
         Height          =   30
         Index           =   0
         Left            =   450
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         Stretch         =   -1  'True
         Top             =   2295
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lbuse 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "心在流浪"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   7
         Top             =   1680
         Width           =   720
      End
      Begin VB.Image uselogo 
         Height          =   1500
         Left            =   120
         OLEDropMode     =   1  'Manual
         Stretch         =   -1  'True
         Top             =   75
         Width           =   1500
      End
      Begin VB.Image E2 
         Height          =   30
         Index           =   2
         Left            =   0
         OLEDropMode     =   1  'Manual
         Stretch         =   -1  'True
         Top             =   2280
         Visible         =   0   'False
         Width           =   5100
      End
      Begin VB.Label LBSG 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Zero To a Hero"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   16
         Top             =   2040
         Width           =   1710
      End
      Begin VB.Image K 
         Height          =   735
         Index           =   0
         Left            =   0
         Top             =   2310
         Width           =   5055
      End
      Begin VB.Shape SB 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   735
         Index           =   3
         Left            =   0
         Top             =   2310
         Width           =   5340
      End
   End
End
Attribute VB_Name = "frmma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'V1.24
'删除多边形边框,方方正正典雅美观
'作者:李双林
'2010年10月起
Option Explicit
Dim sBlendFunction As BLENDFUNCTION '图像渐入渐出の变量
Dim Wnd As Long     '定义显示/隐藏桌面图标的变量
Dim SX As Integer   '定义这个变量是取得鼠标坐标
Dim SY As Integer   '定义这个变量是取得鼠标坐标
Dim pos As POINTAPI '定义这个变量是取得鼠标坐标
Dim hTaskBar As Long '定义这个变量是获得任务栏高度
Dim strpath As String, P_BK_INDEX As Integer
Dim CLOUD As Boolean
Dim LES
Dim ISONME As Boolean, LAST_URL As String
Private PL_NAME As String, HD_NAME As String
Dim IS_PICSHOW As Boolean
Dim ISMOVE As Boolean
Dim SELFW As Boolean
Dim MB As Long
Dim dX, dY
Dim SET_MOVE As Boolean, NOT_M As Boolean, ZOOM_M As Boolean, ZOOM_IN_M As Boolean, ZOOM_OUT_M As Boolean
Dim Ms As String, Ks As Integer, rEADY As Integer, Connect As Integer, THIS_DAY As String
Public MFILEPATH As String, FILETIME As String, FILESINGER As String, SINGERLOGO As String, SINGERLOGOB As String
Dim CLIPS As Collection, iCount As Integer, CTL_MOVE As Boolean
Private m_xPos As Long '上次的鼠标轨迹水平坐标
Private m_yPos As Long '上次的鼠标轨迹垂直坐标
Private m_GeCol As Collection   '记录轨迹点的集合
Dim USEBACK As String
Private gX As Long
Dim TESTMODE As Boolean
Dim MouseDown As Boolean
Dim UNCOUNT As Boolean
Dim MINSOUND As Boolean
Dim MUSIC_MOVE As Boolean
Private mmp As clsNet
Private myIp As String
Private netIp As String
Public AUTOPLAYPIC As Boolean
Const URLTMP = "http://www.baidu.com/baidu.word="
Dim sk As Long, HQ As Long, Counter As Long '为CPU使用率做坐标
Dim once As Boolean '音乐播放器是否启动后第一次使用
Dim pA, fn, Root As String '添加本地驱动器
Const pwdChar = "·" '类化密码框的文字
Dim Pwd As String, PwdLen As Long '密码长度与密码值
Dim SelPos As Long, Insert As Integer
Dim SCREENSAVER As Integer '定义是否运行屏保
Dim TimerStr(0 To 5) As Integer '显示播放事件的变量
Dim sURL As String '定义自动播放曲目的地址
Dim sIndex As Integer '
Dim MovY As Integer '用于进度条
Dim songinx As Integer '判断是否连接超时
Dim imagepixels(2, 1024, 1024) As Integer
Dim Px As Integer
Dim py As Integer
Dim nx As Integer
Dim ny As Integer
Dim a As Integer

Dim imgPlayer As String '下面几行是定义图片--▕
Dim imgPic As String
Dim imgMaster As String
Dim imgBro As String  '      ▕
Dim imgSet As String
Dim imgManager As String
Dim imgFile As String '到这里 ---------------▕

Const HWND_TOPMOST = -1 '置顶
Private Type TGUID
Data1 As Long
data2 As Integer
Data3 As Integer
Data4(0 To 7) As Byte
End Type
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As udtCHOOSECOLOR) As Long
Private Type udtCHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Const CC_FULLOPEN = &H2
Private Const WM_SETHOTKEY = &H32 '定义快捷键
Private Const HOTKEYF_SHIFT = &H1 '定义快捷键shift
Private Const HOTKEYF_CONTROL = &H2       '定义快捷键ctrl
Private Const HOTKEYF_ALT = &H4   '定义快捷键alt
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function PdhVbOpenQuery Lib "PDH.DLL" (ByRef QueryHandle As Long) As Long
Private Declare Function PdhVbAddCounter Lib "PDH.DLL" (ByVal QueryHandle As Long, ByVal CounterPath As String, ByRef CounterHandle As Long) As Long
Private Declare Function PdhCollectQueryData Lib "PDH.DLL" (ByVal QueryHandle As Long) As Long
Private Declare Function PdhVbGetDoubleCounterValue Lib "PDH.DLL" (ByVal CounterHandle As Long, ByRef CounterStatus As Long) As Double
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function OleLoadPicturePath Lib "oleaut32.dll" (ByVal szURLorPath As Long, ByVal punkCaller As Long, ByVal dwReserved As Long, ByVal clrReserved As OLE_COLOR, ByRef riid As TGUID, ByRef ppvRet As IPicture) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public LOADTIME As Long
Private memory As MEMORYSTATUS
Private WithEvents PSubClass As cSubclass '移动窗体无拖影的事件
Attribute PSubClass.VB_VarHelpID = -1

Public Function LoadNetPicture(ByVal ImgSrc As String) As PICTURE
    Dim riid As TGUID
    riid.Data1 = &H7BF80980: riid.data2 = &HBF32: riid.Data3 = &H101A
    riid.Data4(0) = &H8B: riid.Data4(1) = &HBB: riid.Data4(2) = &H0
    riid.Data4(3) = &HAA: riid.Data4(4) = &H0:   riid.Data4(5) = &H30
    riid.Data4(6) = &HC:   riid.Data4(7) = &HAB
    OleLoadPicturePath StrPtr(ImgSrc), 0&, 0&, 0&, riid, LoadNetPicture
End Function
Public Function NTDomainUserName() As String
On Error Resume Next
Dim strBuffer As String * 255
Dim lngBufferLength As Long
Dim lngRet As Long
Dim strTemp As String
lngBufferLength = 255
lngRet = GetUserName(strBuffer, lngBufferLength)
strTemp = UCase(Trim$(strBuffer))
NTDomainUserName = Left$(strTemp, Len(strTemp) - 1)
End Function
Private Sub BACKME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If UNME.PICTURE <> Frmm.PIC(177).PICTURE Then UNME.PICTURE = Frmm.PIC(177).PICTURE
If BACKME.PICTURE <> Frmm.PIC(176).PICTURE Then BACKME.PICTURE = Frmm.PIC(176).PICTURE
If MINIME.PICTURE <> Frmm.PIC(179).PICTURE Then MINIME.PICTURE = Frmm.PIC(179).PICTURE
If SETME.PICTURE <> Frmm.PIC(173).PICTURE Then SETME.PICTURE = Frmm.PIC(173).PICTURE
End Sub

Private Sub BACKME_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
BACKME.Visible = False
PP.Visible = False
If pl.Left = 0 Then
If AUTOPLAYPIC = True Then Timers.Enabled = True
PF(6).Visible = True
pl.AutoRedraw = False
End If
End Sub
Sub SERCHNET()
On Error Resume Next
    Dim I As Integer
    Dim s As String
    Dim f As Integer, t As Integer
    If (IsNumeric(txtFrom.Text)) Then Let f = txtFrom.Text
    If (IsNumeric(txtTo.Text)) Then Let t = txtTo.Text
    If ((f < 0) Or (t < 0)) Then Exit Sub
    If (t < f) Then Exit Sub
    For I = f To t
        lstRes.Enabled = False
        IMJ.Visible = False
        PR.Visible = True
        PR.ZOrder 0
        PR.Width = 1
        PR.Width = (PICTOOL.ScaleWidth / (t / I))
        DoEvents
        'Debug.Print "正在分解主机名..." & S & " " & Int(100 / (T / i)) & "%"
        Let s = netIp & I
        DoEvents
        If (s <> myIp) Then
        DoEvents
            If (mmp.Ping(s)) Then Call Add_Mach(mmp.ResolveHostname(s), s)
        End If
        DoEvents
    Next I
    Call SHOWWRONG("局域网扫描完毕.共发现" & lstRes.ListCount & "台主机", 1)
    IMJ.Visible = True
    lstRes.Enabled = True
    PR.Visible = False
End Sub
Sub ShowFriend()
PICSER.Visible = False
Timers.Enabled = False
Timefriend.Enabled = True
PF(6).Enabled = False
IMPIC.PICTURE = Frmm.PIC(93).PICTURE
IMMAIN.PICTURE = Frmm.PIC(92).PICTURE
IMCHAT.PICTURE = Frmm.PIC(96).image
End Sub
Sub ShowIM()
PICSER.Visible = False
Call SUBDRAWIM
TMRZ.Enabled = True
Timers.Enabled = False
PF(6).Enabled = False
IMPIC.PICTURE = Frmm.PIC(93).PICTURE
IMMAIN.PICTURE = Frmm.PIC(90).PICTURE
IMCHAT.PICTURE = Frmm.PIC(98).image
End Sub
Sub ShowPic()
PICSER.Visible = False
PF(3).ScaleMode = 1
Timetool.Enabled = True
PF(6).Enabled = True
pl.AutoRedraw = False
IMPIC.PICTURE = Frmm.PIC(95).image
IMMAIN.PICTURE = Frmm.PIC(90).image
IMCHAT.PICTURE = Frmm.PIC(96).image

End Sub
Sub LoadParam()
On Error Resume Next '加载用户头像

LOGO = GetSetting("ICEE", "Main", "logo", App.Path + "\Skin\DefaultHead.Bmp")
If Len(Trim(LOGO)) > 0 And PathFileExists(LOGO) <> 0 Then
Picture1.PICTURE = LoadPicture(LOGO)
Else
Call SaveSetting("ICEE", "Main", "logo", App.Path + "\Skin\DefaultHead.Bmp")
Picture1.PICTURE = LoadPicture(App.Path + "\Skin\DefaultHead.Bmp")
End If

Picture2.Height = Picture1.Height
Picture2.Width = Picture1.Width

Call DRAWFACE

PDB.PaintPicture USELOGO.PICTURE, 269, 74, 60, 60
End Sub
Sub setDefaultConfig() '获得电脑名称
Dim computerName As String, strLength As Long
strLength = 255
computerName = String(strLength, Chr(0))
GetComputerName computerName, strLength
LBUSE.Caption = computerName
End Sub
Sub READPASS()
On Error Resume Next
Dim temp As String
temp = GetInitEntry("IM", "LastPassWord", "")
If TXTPOUP.Text <> temp Then
Call SHOWWRONG("密码错误,请重新输入", 0)
TXTPOUP.Text = ""
Else
PF(12).Visible = False
Call LOCKSAFE
TXTPOUP.Text = ""
IS_LOCK = False
End If
End Sub
Sub 导出列表()
On Error Resume Next
Dim sFile As String
sFile = ShowSave(Me.hwnd, "播放列表(*.M3U)" & Chr$(0) & "*.M3U", "导出播放列表")
   Open sFile For Output As #1
    For I = 0 To PLIST.ListCount - 1
     Print #1, PLIST.URL(I)
    Next I
Close #1
End Sub
Sub 打开列表()
Dim sFile As String
sFile = ShowOpen(Me.hwnd, "播放列表(*.M3U)" & Chr$(0) & "*.M3U", "导入播放列表")
If Dir$(sFile) <> vbNullString Then
 Me.Playlist (sFile)
End If
End Sub
Private Sub CZ_COLOR_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
lRet = SetInitEntry("DiskTip", "COLOR", CZ_COLOR(Index).BackColor)
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">更换了界面颜色:" & CZ_COLOR(Index).BackColor
Call DRAWUI
Call DRAWFACE
End Sub

Private Sub E2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 2
If Wm.playState = wmppsPlaying Then
Wm.Controls.currentPosition = X * Wm.currentMedia.duration / (E2(2).Width * 15)
Else
Call UpNow
End If
End Select
End Sub

Private Sub E2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW

End Sub

Private Sub EI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call UpNow
End Sub
Sub SHOWU()
On Error Resume Next
Dim LnBlendPtr As Long
pl.AutoRedraw = False
pl.Cls
For I = 0 To 50
sBlendFunction.SourceConstantAlpha = I * 5
CopyMemory LnBlendPtr, sBlendFunction, 4
AlphaBlend pl.hdc, 0, 0, pl.ScaleWidth, pl.ScaleHeight, IPLAY(MB).MY_PIC.hdc, 0, 0, IPLAY(MB).MY_PIC.Width, IPLAY(MB).MY_PIC.Height, LnBlendPtr
If IS_PICSHOW = False Then
pl.Line (0, 0)-(pl.ScaleWidth, 40), Frmm.PTCO.POINT(0, 0), BF
Call PaintPng(App.Path + "\Skin\UI_TIT.png", pl.hdc, 0, 0) '重绘个性相册标题
Call PaintPng(App.Path + "\SKIN\PO_T.PNG", pl.hdc, IMPIC.Left + 4, 8)
End If
DoEvents
Delay 50
Call PaintPng(App.Path + "\SKIN\PT_S.PNG", pl.hdc, 0, 0)
Next
MB = MB + 1
If MB = filHidden.ListCount Then MB = 0
End Sub
Private Function MyHotKey(vKeyCode) As Boolean
MyHotKey = (GetAsyncKeyState(vKeyCode) < 0)
End Function

Private Sub EI_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub FILHIDDEN_PathChange()
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">更换了个性相册路径:" & filHidden.Path
End Sub

Private Sub Form_Initialize()
SET_MOVE = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Call IFRAME_KeyDown(KeyCode, 0)
End Sub
Private Sub ICC_Click(Index As Integer)
On Error Resume Next
Dim Tfile As String
Tfile = App.Path & "\MEDIA\THUMBS.Bmp"
Select Case Index
Case 0
Call 保存一下(ImgPreview)
Case 1
Call SavePicture(ImgPreview.PICTURE, Tfile)
FRMBOARD.OpenFile (Tfile)
FRMBOARD.Show
Kill Tfile
Case 2
Call PRINTPIC
Case 3
DefCOM = 3
Call SavePicture(ImgPreview.PICTURE, Tfile)
Call SHAREIT(Tfile)
End Select
End Sub

Private Sub ICK_Click(Index As Integer)
Select Case Index
Case 0
lRet = SetInitEntry("IM", "UseNewUser", Not ICK(0).Value)
Case 1

Case 2
lRet = SetInitEntry("IM", "AutoLogin", ICK(2).Value)
End Select
End Sub

Private Sub ICL_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
Call READPASS
Case 2
If Status.RasConnState <> &H2000 Then Call SHOWWRONG("没有检查到网络连接", 2): Exit Sub
Frmadd.Text1.Text = "http://xiazai.xiazaiba.com/Soft/I/IE8-Setup-Ylmf.exe"
Call Frmadd.DOWNLOADIT
PF(2).Visible = False
Case 3
PF(2).Visible = False
lRet = SetInitEntry("SYSTEM", "IETIP", "1")
Case 5
Dim BFPATH As String
BFPATH = BrowseFolder("浏览文件夹", Me)
If BFPATH = "" Then Exit Sub
filHidden.Path = BFPATH '将文件列表设为选中的文件夹
If filHidden.ListCount >= 250 Then Call SHOWWRONG("我嘞个去,您这是想累死我吗.这个文件夹的图片太多了,请重新选择!", 2)
If filHidden.ListCount < 2 Then
filHidden.Path = GetInitEntry("PHOTO_PLAYER", "PATH", App.Path & "\SKIN\PHOTO") '如果那个列表不符合要求当然要回归上一次有效的文件夹
Call SHOWWRONG("个性相册文件夹需要至少2张图片,你选择的文件夹不符合要求,请重新选择", 2): Exit Sub '要退出这个循环
Else
Call LoadPic '合格的话就加载控件并保存路径
lRet = SetInitEntry("PHOTO_PLAYER", "PATH", BFPATH)
End If
Case 6
If ICL(6).MY_TIT = "取消" Then Running = False
If PICLO.Visible = True Then CancelLogin
Case 7
Call SERCHNET
Case 1
Call LogIn
Case 10
If PF(9).Visible = False Then
PF(9).Visible = True
Else
PF(9).Visible = False
End If
Case 9
If txtEntry.Text = "" Then Exit Sub
Call CalculateEntry
txtEntry.Text = ""
Case 12
Call 建议
Case 11
Call BUG
Case 13
Call ADDICQ
Case 14
Call REMOVE
Case 15
Call 修改密码
Case 16
PICPASS.Visible = False: LOCKSAFE: LA(1).Caption = "好友列表": IMJ.Visible = False
Call SUBDRAWIM
Case 8
Call ADDFRIEND
Case 4
Unload Me
End Select
End Sub

Private Sub ICLOS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ICLOS.PICTURE <> Frmm.PIC(178).PICTURE Then ICLOS.PICTURE = Frmm.PIC(178).PICTURE

End Sub

Private Sub ICLOS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PF(17).Visible = False
End Sub

Private Sub ICOCO_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
For I = 0 To ICOCO.Count - 1
ICOCO(I).IS_SELECT = False
Next
ICOCO(Index).IS_SELECT = True
Select Case Index
Case 0
PF(4).BackColor = vbBlack
lRet = SetInitEntry("SYSTEM", "WIN_COLOR", vbBlack)
Case 1
PF(4).BackColor = vbWhite
lRet = SetInitEntry("SYSTEM", "WIN_COLOR", vbWhite)
Case 2
LES = BitBlt(PF(4).hdc, 0, 0, PF(4).Width, PF(4).Height, iFrame.hdc, PF(4).Left, PF(4).Top, &HCC0020)
PF(4).Refresh
End Select
lRet = SetInitEntry("SYSTEM", "WIN_COLOR_SEL", Index)
End Sub

Private Sub ICOCO_MOUSEUP(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If PF(17).Visible = True Then PF(17).Visible = False
End Sub

Private Sub ICP_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Select Case Index
Case 0
Me.PopupMenu Frmm.播放控制, , ICP(0).Left + Pmusic.Left, ICP(0).Top + Mbar.Top + Pmusic.Top + PP.Top + 30
Case 1
Call Lmenu(3)
Case 2
If Pser.Visible = True Then
Pser.Visible = False
PLIST.Move 0, 40, Pmusic.ScaleWidth, Pmusic.ScaleHeight - Mbar.Height - PLIST.Top
Else
Pser.Visible = True
PLIST.Move 0, 40, Pmusic.ScaleWidth, Pmusic.ScaleHeight - Mbar.Height - Pser.Height - PLIST.Top
End If
Case 3
Me.PopupMenu Frmm.顺序, , ICP(3).Left + Pmusic.Left, ICP(3).Top + Mbar.Top + Pmusic.Top + PP.Top + 30
End Select
End Sub

Private Sub ICS_Click(Index As Integer)
Select Case Index
Case 0
PF(5).Visible = True
PF(7).Visible = False
Case 1
PF(5).Visible = False
PF(7).Visible = True
End Select
Dim I As Integer
For I = 0 To ICS.Count - 1
ICS(I).IS_SELECT = False
Next
ICS(Index).IS_SELECT = True
End Sub

Private Sub ICT_CLICK()
Call OpenFile(App.Path & "\COFING\CLIPTEXT.txt")
End Sub



Private Sub ICZ_CLICK(Index As Integer)
On Error Resume Next
Select Case Index
Case 3
PF(3).Move 0, 185, 340, 370
pl.Cls
pl.Move 0, 0, PF(3).ScaleWidth, PF(3).ScaleHeight
ICZ(3).Visible = False
PICTOOL.ZOrder 0
IMJ.ZOrder 0
IS_PICSHOW = False
End Select
End Sub

Private Sub iFrame_DblClick()

If UNME.Enabled = True Then Call NoNoNo
End Sub

Private Sub IFRAME_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape And PICCLIP.Visible = False And PicZoom.Visible = False Then Call NoNoNo
End Sub
Sub GetDriver()
If GetUdisk() <> "" Then
Call KillAuto(GetUdisk)
IMGUSB.ToolTipText = "发现可移动设备"
HASUSB = True
If PICBACK.Visible = True Then PSEND.Visible = True
If PSEND.Visible = True Then
PMDL.Left = PSEND.Left + PSEND.Width + 5
Else
PMDL.Left = PSEND.Left
End If
IMGUSB.PICTURE = Frmm.PIC(38).PICTURE
If Sound = 1 Then sndPlaySound App.Path + "\Sound\popo.wav", 1
Call DRAWMUSIC
End If
End Sub
Sub iCan()
On Error Resume Next
frmma.Show

Dim PBOX As Control
For Each PBOX In Me.Controls
If TypeOf PBOX Is PictureBox Then PBOX.Refresh
Next

Call MoveWindow(frmmabk.hwnd, Me.Left / Screen.TwipsPerPixelX - 20, Me.Top / Screen.TwipsPerPixelY - 10, 380, 650, True)
If ALWAYSONTOP = True Then RESL = SetWindowPos(frmmabk.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags) Else RESL = SetWindowPos(frmmabk.hwnd, 1, 0, 0, 0, 0, flags)
If ALWAYSONTOP = True Then RESL = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags) Else RESL = SetWindowPos(Me.hwnd, 1, 0, 0, 0, 0, flags)
If PICAD.Visible = False Then Frmm.TimeHon.Enabled = True
If PP.Visible = False And pl.Left = 0 And AUTOPLAYPIC = True Then Timers.Enabled = True
If PICD.Visible = True Or PP.Visible = True Then Timers.Enabled = False
Unload FRMTASK
frmmabk.Visible = True
lRet = SetInitEntry("MsgBOX", "LEFT", frmma.Left - 100)
lRet = SetInitEntry("MsgBOX", "TOP", frmma.Top + (frmma.Height - 4000) / 2)
End Sub
Private Sub IFRAME_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Call UpNow
End Sub
Private Sub IFRAME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MOVENOW
If Pmusic.Visible = True Then Pmusic.Visible = False
End Sub
Sub SHOWBK() '显示阴影
On Error Resume Next
frmmabk.Move Me.Left - 20, Me.Top - 10
frmmabk.Visible = True
Me.ZOrder 0
End Sub
Sub HIDEBK()
frmmabk.Hide
End Sub
Private Sub iFrame_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Tit_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub
Private Sub IMAD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If IMAD.PICTURE <> Frmm.PIC(119).image Then IMAD.PICTURE = Frmm.PIC(119).image
End Sub

Private Sub IMAD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMAD.PICTURE <> Frmm.PIC(40).image Then IMAD.PICTURE = Frmm.PIC(40).image
End Sub

Private Sub IMAD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PICAD.ZOrder 0
PICAD.Visible = True
Frmm.TMA.Enabled = True
End Sub



Private Sub IMBAIDU_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If IMBAIDU.PICTURE = Frmm.X2.PICTURE Then IMBAIDU.PICTURE = Frmm.X3.PICTURE
End Sub

Private Sub IMBAIDU_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMBAIDU.PICTURE = Frmm.X1.PICTURE Then IMBAIDU.PICTURE = Frmm.X2.PICTURE
End Sub

Private Sub IMBAIDU_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMBAIDU.PICTURE = Frmm.X3.PICTURE Then IMBAIDU.PICTURE = Frmm.X1.PICTURE: PICSER.Visible = False
End Sub
Private Sub IMCHAT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If IMCHAT.PICTURE <> Frmm.PIC(98).image Then IMCHAT.PICTURE = Frmm.PIC(98).image
End Sub

Private Sub IMCHAT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMAD.PICTURE <> Frmm.PIC(35).image Then IMAD.PICTURE = Frmm.PIC(35).image
IMAD.Visible = False
If IMCHAT.PICTURE <> Frmm.PIC(97).image Then IMCHAT.PICTURE = Frmm.PIC(97).PICTURE
End Sub

Private Sub IMCHAT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If PICIM.Left = 0 Then Exit Sub
Call ShowIM
End Sub

Private Sub IMCLEAR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CTL_MOVE = False Then
CTL_MOVE = True
LES = BitBlt(IMCLEAR.hdc, 0, 0, IMCLEAR.Width, IMCLEAR.Height, PP.hdc, IMCLEAR.Left, IMCLEAR.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\DE_H.PNG", IMCLEAR.hdc, 0, 0)
IMCLEAR.Refresh
End If
End Sub

Private Sub IMCLEAR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
On Error Resume Next
fso.DeleteFile SINGERLOGO
Call GETSINGER
End Sub

Private Sub IMCLIP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Sleep 100
Call NoNoNo
Capture.Show
End Sub

Private Sub IMCLIP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMCLIP.PICTURE <> Frmm.PIC(26).PICTURE Then IMCLIP.PICTURE = Frmm.PIC(26).PICTURE
End Sub

Private Sub IMCLP_Click()
If IS_CHECK_CLIP = True Then
StopMonitoring Me.hwnd '停止
IS_CHECK_CLIP = False
IMCLP.PICTURE = Frmm.PIC(131).image
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">停止了剪切板监视"
Else
IMCLP.PICTURE = Frmm.PIC(130).image
StartMonitoring Me.hwnd '开始
IS_CHECK_CLIP = True
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">开始了剪切板监视"
End If
lRet = SetInitEntry("SYSTEM", "CHECK_CLIP", IS_CHECK_CLIP)

End Sub

Private Sub IMCPU_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub IMCPU_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PICCPU.Visible = False Then PICCPU.Visible = True: PICCPU.ZOrder 0: PICCPU.Move 8, 216
End Sub

Private Sub IMCZ_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Select Case Index
Case 0
LA(4).Caption = LA(4).Caption + 100
Case 1
If LA(4).Caption = 100 Then Exit Sub
LA(4).Caption = LA(4).Caption - 100
End Select
mfScale = LA(4).Caption / 100!
lRet = SetInitEntry("SCREEN_MAKER", "ZOOMPER", LA(4).Caption)
End Sub

Private Sub IMCZ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
If ZOOM_IN_M = False Then ZOOM_IN_M = True
Case 1
If ZOOM_OUT_M = False Then ZOOM_OUT_M = True
End Select
End Sub

Private Sub IMEND_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If IMEND.PICTURE = Frmm.X2.PICTURE Then IMEND.PICTURE = Frmm.X3.PICTURE

End Sub

Private Sub IMEND_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMEND.PICTURE = Frmm.X1.PICTURE Then IMEND.PICTURE = Frmm.X2.PICTURE
End Sub

Private Sub IMEND_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMEND.PICTURE = Frmm.X3.PICTURE Then IMEND.PICTURE = Frmm.X1.PICTURE
PF(2).Visible = False
End Sub


Private Sub IMG_NT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PICNET.Visible = True Then Exit Sub
If SET_MOVE = False Then
SET_MOVE = True
PICLO.Cls
LES = BitBlt(PICLO.hdc, 0, 0, PICLO.Width, PICLO.Height, PF(3).hdc, PICLO.Left, PICLO.Top, &HCC0020)
PICLO.Line (43, 88)-(303, 330), iFrame.BackColor, BF

Call PaintPng(App.Path + "\Skin\login.png", PICLO.hdc, 0, 0) '登陆界面
PICLO.Line (0, 0)-(PICLO.ScaleWidth, 40), Frmm.PTCO.POINT(0, 0), BF
Call PaintPng(App.Path + "\Skin\UI_TIT.png", PICLO.hdc, 0, 0) '重绘登陆框标题
Call PaintPng(App.Path + "\SKIN\PO_T.PNG", PICLO.hdc, IMCHAT.Left + 4, 8)
Call PaintPng(App.Path & "\SKIN\SET_H.PNG", PICLO.hdc, 8, 0)
PICLO.Refresh

If IMSER.Visible = True Then IMSER.Visible = False
If IMSER.PICTURE <> Frmm.PIC(81).PICTURE Then IMSER.PICTURE = Frmm.PIC(81).PICTURE
If IMMAIN.PICTURE <> Frmm.PIC(90).PICTURE Then IMMAIN.PICTURE = Frmm.PIC(90).PICTURE

End If
End Sub

Private Sub IMG_NT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If PICNET.Visible = True Then Exit Sub
PICLO.Cls
PICLO.BackColor = Frmm.PTCO.POINT(0, 0)
IMG_NT.Visible = False
PICNET.Visible = True
IMJ.Visible = True
LBITEM(2).Caption = "扫描局域网"
Call DRAWNET
Call RUNSAFE
PICLO.Cls
PICLO.BackColor = Frmm.PTCO.POINT(0, 0)
End Sub

Private Sub IMGEM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMSKIN.PICTURE <> Frmm.PIC(12).PICTURE Then IMSKIN.PICTURE = Frmm.PIC(12).PICTURE
If IMGEM.PICTURE = Frmm.PIC(68).PICTURE Then IMGEM.PICTURE = Frmm.PIC(17).PICTURE
End Sub
Private Sub IMGEM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMGEM.PICTURE = Frmm.PIC(17).PICTURE Then IMGEM.PICTURE = Frmm.PIC(68).PICTURE
If Button <> 1 Then Exit Sub
If frmma.Left > FRMHIS.Width Then
FRMHIS.Move frmma.Left - FRMHIS.Width, frmma.Top
Else
FRMHIS.Move frmma.Left + frmma.Width, frmma.Top
End If
FRMHIS.Show
GETMSGCOUNT = 0
LCO.Caption = 0
End Sub

Private Sub IMGFAV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Wm.URL = "" Then Exit Sub
 If FAV_IT = False Then
Call FRMFAV.ADD_ITEM(SONGNAME, Wm.URL)
FAV_IT = True
IMGFAV.PICTURE = Frmm.PIC(52).PICTURE
If IS_NET = True Then FrmNetMusic.IMFAV.PICTURE = Frmm.PIC(52).PICTURE
Else
If IS_NET = True Then FrmNetMusic.IMFAV.PICTURE = Frmm.PIC(54).PICTURE
FAV_IT = False
IMGFAV.PICTURE = Frmm.PIC(54).PICTURE
Call FRMFAV.REMOVE_ITEM(SONGNAME)
End If
End Sub

Private Sub IMGUSB_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Tit_OLEDragDrop(Data, Effect, Button, Shift, X, Y)

End Sub

Private Sub IMKK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If Wm.URL = "" Then Exit Sub
If Wm.playState = wmppsPlaying Then
Wm.Controls.pause
TMP.Enabled = False
IW(0).SETPNG App.Path & "\SKIN\P_N.png", 70, 70
Else
Wm.Controls.Play
TMP.Enabled = True
IW(0).SETPNG App.Path & "\SKIN\PA_N.png", 70, 70
End If
End Sub

Private Sub IMMAIN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If IMMAIN.PICTURE <> Frmm.PIC(92).PICTURE Then IMMAIN.PICTURE = Frmm.PIC(92).PICTURE
End Sub

Private Sub IMMAIN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMMAIN.PICTURE <> Frmm.PIC(91).PICTURE Then IMMAIN.PICTURE = Frmm.PIC(91).PICTURE
End Sub

Private Sub IMMAIN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If PicUse.Left = 0 Then Exit Sub
Call ShowFriend
End Sub


Private Sub IMPIC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If IMPIC.PICTURE <> Frmm.PIC(95).image Then IMPIC.PICTURE = Frmm.PIC(95).image
End Sub

Private Sub IMPIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMPIC.PICTURE <> Frmm.PIC(94).image Then IMPIC.PICTURE = Frmm.PIC(94).image
End Sub

Private Sub IMPIC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If pl.Left = 0 Then Exit Sub
Call ShowPic
End Sub

Private Sub IMR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If AUTOPLAYPIC = False Then
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">开启了相册播放"
AUTOPLAYPIC = True
IMR.PICTURE = Frmm.PIC(130).image
Else
IMR.PICTURE = Frmm.PIC(131).image
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">关闭了相册播放"
AUTOPLAYPIC = False
End If
lRet = SetInitEntry("SYSTEM", "AUTOPLAYPIC", AUTOPLAYPIC)
End Sub
Private Sub IMSER_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If PICSER.Visible = False Then IMSER.PICTURE = Frmm.PIC(83).image
If PICSER.Visible = True Then Call UpNow
End Sub

Private Sub IMSER_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PICSER.Visible = False Then IMSER.PICTURE = Frmm.PIC(82).image
End Sub

Private Sub IMSER_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PICSER.Visible = True Then Exit Sub
PICSER.Visible = True
TXTBAIDU.Text = "<请输入关键词>"
TXTBAIDU.SetFocus
End Sub
Sub 搜索封面()
Call Frmm.CHECKNET
If Status.RasConnState <> &H2000 Then Exit Sub
If Trim(FILESINGER) = "" Then Exit Sub
Call FRMFM.SERCHSINGER(FILESINGER)
End Sub

Private Sub IMSERG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CTL_MOVE = False Then
CTL_MOVE = True
LES = BitBlt(IMSERG.hdc, 0, 0, IMSERG.Width, IMSERG.Height, PP.hdc, IMSERG.Left, IMSERG.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\SS_H.PNG", IMSERG.hdc, 0, 0)
IMSERG.Refresh
End If
End Sub

Private Sub IMSERG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Call 搜索封面
End Sub

Private Sub IMSERG_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PP_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub


Private Sub IMSIN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub IMSIN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PF(8).Visible = False Then PF(8).Visible = True: PF(8).ZOrder 0: PF(8).Move 176, 515
End Sub

Private Sub IMSKIN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If IMSKIN.PICTURE <> Frmm.PIC(12).PICTURE Then IMSKIN.PICTURE = Frmm.PIC(12).PICTURE

End Sub

Private Sub IMSKIN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMGEM.PICTURE <> Frmm.PIC(68).PICTURE Then IMGEM.PICTURE = Frmm.PIC(68).PICTURE
If IMSKIN.PICTURE <> Frmm.PIC(13).PICTURE Then IMSKIN.PICTURE = Frmm.PIC(13).PICTURE

End Sub

Private Sub IMSKIN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
If PF(15).Visible = False Then
PF(15).Visible = True
PF(15).ZOrder 0
PF(15).SetFocus
Else
PF(15).Visible = False
End If
End Sub
Private Sub IMGUSB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMGUSB.PICTURE <> Frmm.PIC(38).PICTURE Then Exit Sub
If Button = 1 Then Call SafeUdisk(GetUdisk())
End Sub

Private Sub IMGUSB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMCLIP.PICTURE <> Frmm.PIC(24).PICTURE Then IMCLIP.PICTURE = Frmm.PIC(24).PICTURE
If PICCPU.Visible = True Then PICCPU.Visible = False
End Sub
Private Sub IMJ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMJ.PICTURE = Frmm.X2.PICTURE And Button = 1 Then IMJ.PICTURE = Frmm.X3.PICTURE
End Sub

Private Sub IMJ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMJ.PICTURE = Frmm.X1.PICTURE Then IMJ.PICTURE = Frmm.X2.PICTURE
End Sub

Private Sub IMJ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If IMJ.PICTURE = Frmm.X3.PICTURE Then IMJ.PICTURE = Frmm.X1.PICTURE '首先初始化图片
If MAINSTYLE = 3 Then PF(4).Visible = True
USEBACK = GetInitEntry("SYSTEM", "BACKPICTURE", App.Path + "\SKIN\BK\0.JPG") '获得背景图片
Call SUBDRAW

If PICCLIP.Visible = True Then PICCLIP.Visible = False: LBITEM(3).Caption = "主菜单": IMJ.Visible = False: LOCKSAFE: txtText.Visible = False: PVP.Cls: IPR.PICTURE = LoadPicture("") '如果用户正在剪切板界面
If PICD.Visible = True Then Call LEAVEPIC '如果用户正在设置幻灯片
If PICBUG.Visible = True Then PICBUG.Visible = False: LOCKSAFE: LA(1).Caption = "好友列表": IMJ.Visible = False '如果用户正在反馈问题界面
If PICFI.Visible = True Then PICFI.Visible = False: LOCKSAFE: LA(1).Caption = "好友列表": IMJ.Visible = False '如果用户正在查看好友信息
If PICIG.Visible = True Then PICIG.Visible = False: LOCKSAFE: LA(1).Caption = "好友列表": IMJ.Visible = False '如果用户正在黑名单界面
If PICPASS.Visible = True Then PICPASS.Visible = False: LOCKSAFE: LA(1).Caption = "好友列表": IMJ.Visible = False '如果用户正在修改密码
If PICNET.Visible = True Then PICNET.Visible = False: IMG_NT.Visible = True: IMJ.Visible = False: LBITEM(2).Caption = "请登陆": Call LOCKSAFE '如果用户正在扫描局域网

If PF(0).Visible = True Then PF(0).Visible = False: IMJ.Visible = False: Call LOCKSAFE
If IMJ.BackColor <> Frmm.PTCO.POINT(0, 0) Then IMJ.BackColor = Frmm.PTCO.POINT(0, 0)
End Sub
Sub 截屏()
Me.Hide
Call HIDEBK
Capture.Show
End Sub
Private Sub IMVOL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If PLAYB.Visible = True Then PLAYB.Visible = False
If INEXT.Visible = True Then INEXT.Visible = False
If IPRE.Visible = True Then IPRE.Visible = False
If IMVOL.Visible = True Then IMVOL.Visible = False

If CTL_MOVE = False Then
CTL_MOVE = True
PLAYB.Cls
IPRE.Cls
INEXT.Cls
INEXT.BackColor = PP.BackColor
IPRE.BackColor = PP.BackColor
PLAYB.BackColor = PP.BackColor
LES = BitBlt(PLAYB.hdc, 0, 0, PLAYB.Width, PLAYB.Height, PP.hdc, PLAYB.Left, PLAYB.Top, &HCC0020)
If Wm.playState = wmppsPlaying Then Call PaintPng(App.Path & "\SKIN\PA_H.PNG", PLAYB.hdc, 0, 0) Else Call PaintPng(App.Path & "\SKIN\P_H.PNG", PLAYB.hdc, 0, 0)
PLAYB.Refresh
LES = BitBlt(IPRE.hdc, 0, 0, IPRE.Width, IPRE.Height, PP.hdc, IPRE.Left, IPRE.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\PR_N.PNG", IPRE.hdc, 0, 0)
IPRE.Refresh
LES = BitBlt(INEXT.hdc, 0, 0, INEXT.Width, INEXT.Height, PP.hdc, INEXT.Left, INEXT.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\NX_N.PNG", INEXT.hdc, 0, 0)
INEXT.Refresh
End If

If PV.Visible = False Then PV.Visible = True
End Sub
Private Sub INEXT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CTL_MOVE = False Then
CTL_MOVE = True
PLAYB.Cls
IPRE.Cls
INEXT.Cls
INEXT.BackColor = PP.BackColor
IPRE.BackColor = PP.BackColor
PLAYB.BackColor = PP.BackColor
LES = BitBlt(PLAYB.hdc, 0, 0, PLAYB.Width, PLAYB.Height, PP.hdc, PLAYB.Left, PLAYB.Top, &HCC0020)
If Wm.playState = wmppsPlaying Then Call PaintPng(App.Path & "\SKIN\PA_N.PNG", PLAYB.hdc, 0, 0) Else Call PaintPng(App.Path & "\SKIN\P_N.PNG", PLAYB.hdc, 0, 0)
PLAYB.Refresh
LES = BitBlt(IPRE.hdc, 0, 0, IPRE.Width, IPRE.Height, PP.hdc, IPRE.Left, IPRE.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\PR_N.PNG", IPRE.hdc, 0, 0)
IPRE.Refresh
LES = BitBlt(INEXT.hdc, 0, 0, INEXT.Width, INEXT.Height, PP.hdc, INEXT.Left, INEXT.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\NX_H.PNG", INEXT.hdc, 0, 0)
INEXT.Refresh

End If

End Sub

Private Sub INEXT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Select Case LOLIPOP
Case 1, 2, 3
Call NT(2)
Case 0
Call NT(3)
End Select
End Sub

Private Sub INEXT_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub IPRE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call NT(1)
End Sub

Private Sub IPRE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CTL_MOVE = False Then
CTL_MOVE = True
PLAYB.Cls
IPRE.Cls
INEXT.Cls
INEXT.BackColor = PP.BackColor
IPRE.BackColor = PP.BackColor
PLAYB.BackColor = PP.BackColor
LES = BitBlt(PLAYB.hdc, 0, 0, PLAYB.Width, PLAYB.Height, PP.hdc, PLAYB.Left, PLAYB.Top, &HCC0020)
If Wm.playState = wmppsPlaying Then Call PaintPng(App.Path & "\SKIN\PA_N.PNG", PLAYB.hdc, 0, 0) Else Call PaintPng(App.Path & "\SKIN\P_N.PNG", PLAYB.hdc, 0, 0)
PLAYB.Refresh
LES = BitBlt(IPRE.hdc, 0, 0, IPRE.Width, IPRE.Height, PP.hdc, IPRE.Left, IPRE.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\PR_H.PNG", IPRE.hdc, 0, 0)
IPRE.Refresh
LES = BitBlt(INEXT.hdc, 0, 0, INEXT.Width, INEXT.Height, PP.hdc, INEXT.Left, INEXT.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\NX_N.PNG", INEXT.hdc, 0, 0)
INEXT.Refresh
LES = BitBlt(PZOR.hdc, 0, 0, PZOR.Width, PZOR.Height, PP.hdc, PZOR.Left, PZOR.Top, &HCC0020)
If LOLIPOP = 3 Then
Call PaintPng(App.Path & "\SKIN\SX_N.PNG", PZOR.hdc, 0, 0)
ElseIf LOLIPOP = 1 Then
Call PaintPng(App.Path & "\SKIN\DQ_N.PNG", PZOR.hdc, 0, 0)
ElseIf LOLIPOP = 2 Then
Call PaintPng(App.Path & "\SKIN\XH_N.PNG", PZOR.hdc, 0, 0)
ElseIf LOLIPOP = 0 Then
Call PaintPng(App.Path & "\SKIN\SJ_N.PNG", PZOR.hdc, 0, 0)
End If
PZOR.Refresh
End If
If PV.Visible = True Then PV.Visible = False
End Sub

Private Sub IPRE_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub ISHA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CTL_MOVE = False Then
CTL_MOVE = True
LES = BitBlt(ISHA.hdc, 0, 0, ISHA.Width, ISHA.Height, PP.hdc, ISHA.Left, ISHA.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\SHARE_H.PNG", ISHA.hdc, 0, 0)
 ISHA.Refresh
End If

End Sub

Private Sub ISHA_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If PF(13).Visible = False Then PF(13).Visible = True
End Sub

Private Sub IST_CLICK(Index As Integer)
Call SaveSetting("ICEE", "MAIN", "STYLE", Index)
Call ULock
Call OUTPUTTHUMB
Call SUBDRAW
Call DRAWFACE
Call Frmm.LoadStyle
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">更换了主题" & Index
Dim I As Integer
For I = 0 To IST.Count - 1
IST(I).IS_SELECT = False
Next
IST(Index).IS_SELECT = True
End Sub
Private Sub IU_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MUSIC_MOVE = False Then
MUSIC_MOVE = True
IU.Cls
If IU.BackColor <> PP.BackColor Then IU.BackColor = PP.BackColor
LES = BitBlt(IU.hdc, 0, 0, IU.Width, IU.Height, Pmusic.hdc, IU.Left, IU.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\MB_H.PNG", IU.hdc, 5, 3)
IU.Refresh
End If
End Sub
Private Sub IU_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
Pmusic.Visible = False
If Left(UCase(Wm.URL), 7) = "HTTP://" Then PMDL.Visible = True Else PMDL.Visible = False
If HASUSB = True Then PSEND.Visible = True Else PSEND.Visible = False
If PSEND.Visible = True Then PMDL.Left = PSEND.Left + PSEND.Width + 5 Else PMDL.Left = PSEND.Left
lRet = SetInitEntry("PLAYER", "MUSICMODE", 1)
End Sub

Private Sub IU_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub IW_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
Select Case Index
Case 0

Call ShowMusic
Case 1

If frmma.Left > FRMEND.Left Then
FRMEND.Move frmma.Left - FRMEND.Width, frmma.Top
Else
FRMEND.Move frmma.Left + frmma.Width, frmma.Top
End If

FRMEND.Show
Case 2
frmset.Show
Case 3
Call DRAWCLIP
PICCLIP.Visible = True
PF(4).Visible = False
Call RUNSAFE
ListView1.ListItems(ListView1.ListItems.Count).EnsureVisible
ListView1.ListItems(ListView1.ListItems.Count).Selected = True
ListView1.SetFocus
Call SHOWDEMO
LBITEM(3).Caption = "剪切板"
PicUse.BackColor = Frmm.PTCO.POINT(1, 1)
IMJ.BackColor = Frmm.PTCO.POINT(0, 0)
IMJ.Visible = True
Case 4
Call DRAWCAL
Call RUNSAFE
IMJ.Visible = True
PF(0).Visible = True
PF(0).ZOrder 0
PICTOOL.ZOrder 0
IMJ.ZOrder 0
PF(4).Visible = False
Case 5
If frmma.Winsock1.State <> 7 Then
Call SHOWWRONG("请先登录服务器!", 0)
Exit Sub
Else
If frmma.Left > FRMMYID.Left Then
FRMMYID.Move frmma.Left - FRMMYID.Width, frmma.Top
Else
FRMMYID.Move frmma.Left + frmma.Width, frmma.Top
End If
FRMMYID.Show
End If
Case 6
FrmWhatNew.Show
Case 7
Call SHOWDL
Case 8
If Me.Left < FRMDATE.Width / 2 Then
FRMDATE.Move Me.Left + Me.Width, Me.Top
Else
FRMDATE.Move Me.Left - FRMDATE.Width, Me.Top
End If
FRMDATE.Show
Case 9
FRMBOARD.Show
Case 10
Call SHOWZOOM
Case 11

End Select
If PF(17).Visible = True Then PF(17).Visible = False
End Sub

Private Sub IW_MOUSEMOVE(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MOVENOW
End Sub

Private Sub IW_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub IWG_MOUSEMOVE(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SetHand
MOVENOW
End Sub

Private Sub IWG_MOUSEUP(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
ShellExecute 0&, vbNullString, IWG(Index).MYTIT, vbNullString, vbNullString, 0  '调用ie
End Sub

Private Sub IWILLBK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ZOOM_M = False Then ZOOM_M = True
End Sub

Private Sub IWILLBK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
tmrZoom.Enabled = False
PicZoom.Visible = False
Call LOCKSAFE
End Sub

Private Sub JM_COLOR_Click(Index As Integer)
lRet = SetInitEntry("WIN8_DESK", "COLOR", JM_COLOR(Index).BackColor)
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">更换了瓷砖颜色:" & JM_COLOR(Index).BackColor
Call DRAWUI
End Sub

Private Sub K_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Select Case Index
Case 5
If Wm.playState = wmppsPlaying Then
Wm.Controls.currentPosition = X * Wm.currentMedia.duration / (E2(2).Width * 15)
Else
UpNow
End If
End Select
Call UpNow
End Sub
Private Sub k_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMCLIP.PICTURE <> Frmm.PIC(24).PICTURE Then IMCLIP.PICTURE = Frmm.PIC(24).PICTURE
If PICCPU.Visible = True Then PICCPU.Visible = False
If PF(15).Visible = True Then PF(15).Visible = False
If IMSKIN.PICTURE <> Frmm.PIC(12).PICTURE Then IMSKIN.PICTURE = Frmm.PIC(12).PICTURE
If IMAD.PICTURE <> Frmm.PIC(35).image Then IMAD.PICTURE = Frmm.PIC(35).image: IMAD.Visible = False
If PICSER.Visible = False Then
If IMSER.Visible = True Then IMSER.Visible = False
If IMSER.PICTURE <> Frmm.PIC(81).PICTURE Then IMSER.PICTURE = Frmm.PIC(81).PICTURE
Else
If IMSER.Visible = False Then IMSER.Visible = True
If IMSER.PICTURE <> Frmm.PIC(83).PICTURE Then IMSER.PICTURE = Frmm.PIC(83).PICTURE
End If

If MUSIC_MOVE = True Then
MUSIC_MOVE = False
PICBACK.Cls
PICDL.Cls
IU.Cls
If PICBACK.BackColor <> PP.BackColor Then PICBACK.BackColor = PP.BackColor
If IU.BackColor <> PP.BackColor Then IU.BackColor = PP.BackColor
If PICDL.BackColor <> COLOR_NOR Then PICDL.BackColor = COLOR_NOR
LES = BitBlt(PICBACK.hdc, 0, 0, IU.Width, PICBACK.Height, PP.hdc, PICBACK.Left, PICBACK.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\L_N.PNG", PICBACK.hdc, 8, 6)
PICBACK.Refresh
LES = BitBlt(PICDL.hdc, 0, 0, PICDL.Width, PICDL.Height, iFrame.hdc, PICDL.Left, PICDL.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\D_N.PNG", PICDL.hdc, 5, 3)
PICDL.Refresh
LES = BitBlt(IU.hdc, 0, 0, IU.Width, IU.Height, Pmusic.hdc, IU.Left, IU.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\MB_N.PNG", IU.hdc, 5, 3)
IU.Refresh
End If
If SET_MOVE = True And PICNET.Visible = False Then
SET_MOVE = False
PICLO.Cls
LES = BitBlt(PICLO.hdc, 0, 0, PICLO.Width, PICLO.Height, PF(3).hdc, PICLO.Left, PICLO.Top, &HCC0020)
PICLO.Line (43, 88)-(303, 330), iFrame.BackColor, BF
Call PaintPng(App.Path + "\Skin\login.png", PICLO.hdc, 0, 0) '登陆界面
PICLO.Line (0, 0)-(PICLO.ScaleWidth, 40), Frmm.PTCO.POINT(0, 0), BF
Call PaintPng(App.Path + "\Skin\UI_TIT.png", PICLO.hdc, 0, 0) '重绘登陆框标题
Call PaintPng(App.Path + "\SKIN\PO_T.PNG", PICLO.hdc, IMCHAT.Left + 4, 8)
Call PaintPng(App.Path & "\SKIN\SET_N.PNG", PICLO.hdc, 8, 0)
PICLO.Refresh
End If
If PicUse.Left = 0 Then
If IMMAIN.PICTURE <> Frmm.PIC(92).image Then IMMAIN.PICTURE = Frmm.PIC(92).image
If IMPIC.PICTURE <> Frmm.PIC(93).image Then IMPIC.PICTURE = Frmm.PIC(93).image
If IMCHAT.PICTURE <> Frmm.PIC(96).image Then IMCHAT.PICTURE = Frmm.PIC(96).image
ElseIf pl.Left = 0 Then
If IMMAIN.PICTURE <> Frmm.PIC(90).image Then IMMAIN.PICTURE = Frmm.PIC(90).image
If IMPIC.PICTURE <> Frmm.PIC(95).image Then IMPIC.PICTURE = Frmm.PIC(95).image
If IMCHAT.PICTURE <> Frmm.PIC(96).image Then IMCHAT.PICTURE = Frmm.PIC(96).image
ElseIf PICIM.Left = 0 Then
If IMMAIN.PICTURE <> Frmm.PIC(90).image Then IMMAIN.PICTURE = Frmm.PIC(90).image
If IMPIC.PICTURE <> Frmm.PIC(93).image Then IMPIC.PICTURE = Frmm.PIC(93).image
If IMCHAT.PICTURE <> Frmm.PIC(98).image Then IMCHAT.PICTURE = Frmm.PIC(98).image
End If
Select Case Index
Case 0
'If Status.RasConnState <> &H2000 Then Exit Sub '检查Internet
If X >= 30 And X <= 450 Then IMSER.Visible = True
If X >= 4500 And X <= 5000 Then IMAD.Visible = True
End Select
End Sub

Private Sub K_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub LA_Change(Index As Integer)
Select Case Index
Case 3
If LA(3).Caption > 0 Then PLIST.Visible = True Else PLIST.Visible = False
If LA(3).Caption > 999 Then LA(3).Caption = "999"
End Select
End Sub

Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Select Case Index
Case 19, 20
Case Else
Call UpNow
End Select
End Sub
Private Sub LA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMCLIP.PICTURE <> Frmm.PIC(24).PICTURE Then IMCLIP.PICTURE = Frmm.PIC(24).PICTURE
Select Case Index
Case 10
LA(10).ToolTipText = LA(10).Caption
Case 20
If MUSIC_MOVE = True Then
MUSIC_MOVE = False
PICBACK.Cls
PICDL.Cls
IU.Cls
If PICBACK.BackColor <> PP.BackColor Then PICBACK.BackColor = PP.BackColor
If IU.BackColor <> PP.BackColor Then IU.BackColor = PP.BackColor
If PICDL.BackColor <> COLOR_NOR Then PICDL.BackColor = COLOR_NOR
LES = BitBlt(PICBACK.hdc, 0, 0, PICBACK.Width, PICBACK.Height, PP.hdc, PICBACK.Left, PICBACK.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\L_H.PNG", PICBACK.hdc, 8, 6)
PICBACK.Refresh
LES = BitBlt(PICDL.hdc, 0, 0, PICDL.Width, PICDL.Height, iFrame.hdc, PICDL.Left, PICDL.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\D_N.PNG", PICDL.hdc, 5, 3)
PICDL.Refresh
LES = BitBlt(IU.hdc, 0, 0, IU.Width, IU.Height, Pmusic.hdc, IU.Left, IU.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\MB_N.PNG", IU.hdc, 5, 3)
IU.Refresh
End If
Case 19
If MUSIC_MOVE = True Then
MUSIC_MOVE = False
PICBACK.Cls
PICDL.Cls
IU.Cls
If PICBACK.BackColor <> PP.BackColor Then PICBACK.BackColor = PP.BackColor
If IU.BackColor <> PP.BackColor Then IU.BackColor = PP.BackColor
If PICDL.BackColor <> COLOR_NOR Then PICDL.BackColor = COLOR_NOR
LES = BitBlt(PICBACK.hdc, 0, 0, PICBACK.Width, PICBACK.Height, PP.hdc, PICBACK.Left, PICBACK.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\L_N.PNG", PICBACK.hdc, 8, 6)
PICBACK.Refresh
LES = BitBlt(PICDL.hdc, 0, 0, PICDL.Width, PICDL.Height, iFrame.hdc, PICDL.Left, PICDL.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\D_H.PNG", PICDL.hdc, 5, 3)
PICDL.Refresh
LES = BitBlt(IU.hdc, 0, 0, IU.Width, IU.Height, Pmusic.hdc, IU.Left, IU.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\MB_N.PNG", IU.hdc, 5, 3)
IU.Refresh
End If
Case 42, 43, 12, 45, 46, 9, 41, 7, 10, 6, 8

Case Else
Call MOVENOW
End Select
End Sub
Private Sub LA_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Select Case Index
Case 19
Call SHOWDL
Case 20
Call SHOWUIT
Case 21
On Error GoTo ERR:
If PMDL.BackColor <> PP.BackColor Then PMDL.BackColor = COLOR_NOR
Call DoFileDownload(StrConv(Wm.URL, vbUnicode))
End Select
ERR:
Exit Sub

End Sub

Private Sub LBFA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub LBFE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub
Private Sub LBFN_Change()
If LBFN.Height > 36 Then LBFN.Height = 36
End Sub

Private Sub LBFN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpNow
End Sub

Private Sub LBFN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub



Private Sub LBFQ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub


Private Sub LBFS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub
Private Sub LBFW_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub LBITEM_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 3, 2, 1, 12
Call UpNow
End Select
End Sub
Sub RUNSAFE()
IMMAIN.Enabled = False
IMSKIN.Enabled = False
IMPIC.Enabled = False
IMCHAT.Enabled = False
USELOGO.Enabled = False
SETME.Enabled = False
LCO.Enabled = False
PICMU.Enabled = False
UNME.Enabled = False
MINIME.Enabled = False
If pl.Left = 0 Then
Timers.Enabled = False
PF(6).Enabled = False
End If
BACKME.Enabled = False
IMGEM.Enabled = False
IMSER.Enabled = False
PICSER.Visible = False
If PF(15).Visible = True Then PF(15).Visible = False
CAN_SHOW_MEUN = False
frmmp.Hide
K(0).Enabled = False
End Sub
Sub LOCKSAFE()
If PF(15).Visible = True Then PF(15).Visible = False
IMSKIN.Enabled = True
IMMAIN.Enabled = True
IMPIC.Enabled = True
IMCHAT.Enabled = True
UNME.Enabled = True
MINIME.Enabled = True
USELOGO.Enabled = True
PICMU.Enabled = True
SETME.Enabled = True
If pl.Left = 0 Then
If AUTOPLAYPIC = True Then Timers.Enabled = True
PF(6).Enabled = True
End If
BACKME.Enabled = True
IMGEM.Enabled = True
LCO.Enabled = True
PR.Visible = False
IMSER.Enabled = True
CAN_SHOW_MEUN = True
K(0).Enabled = True
End Sub
Private Sub LBITEM_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
On Error Resume Next
Select Case Index
Case 0
If frmma.Winsock1.State <> 7 Then
Call SHOWWRONG("请先登录服务器!", 0)
Exit Sub
Else
If frmma.Left > FRMMYID.Left Then
FRMMYID.Move frmma.Left - FRMMYID.Width, frmma.Top
Else
FRMMYID.Move frmma.Left + frmma.Width, frmma.Top
End If
FRMMYID.Show
End If
Case 4
Call DRAWCLIP
PICCLIP.Visible = True
ListView1.ListItems(ListView1.ListItems.Count).EnsureVisible
ListView1.ListItems(ListView1.ListItems.Count).Selected = True
ListView1.SetFocus
Call SHOWDEMO
Call RUNSAFE
LBITEM(3).Caption = "剪切板"
PicUse.BackColor = Frmm.PTCO.POINT(1, 1)
IMJ.BackColor = Frmm.PTCO.POINT(0, 0)
IMJ.Visible = True
Case 5
frmset.Show
Case 6
FrmWhatNew.Show
Case 8
Call LoadNote
Case 7
Call ShowMusic
Case 10
If frmma.Left > FRMEND.Left Then
FRMEND.Move frmma.Left - FRMEND.Width, frmma.Top
Else
FRMEND.Move frmma.Left + frmma.Width, frmma.Top
End If
FRMEND.Show
Case 11
Call DRAWCAL
Call RUNSAFE
IMJ.Visible = True
PF(0).Visible = True
PF(0).ZOrder 0
PICTOOL.ZOrder 0
IMJ.ZOrder 0
End Select
End If
If PF(17).Visible = True Then PF(17).Visible = False
End Sub

Private Sub LBITEM_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub
Private Sub LBSG_Change()
If Trim(LBSG.Caption) = "" Then
MyPersonalInfo.Country = "这个人很懒，什么都没留下"
LBSG.Caption = MyPersonalInfo.Country
End If
End Sub

Private Sub LBSG_DblClick()
If UNME.Enabled = True Then Call NoNoNo
End Sub

Private Sub LBSG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Call UpNow
End Sub

Private Sub LBSG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LBSG.ToolTipText = Trim(LBSG.Caption)
End Sub

Sub UpNow()
frmmp.Hide
If Pmusic.Visible = True Then Pmusic.Visible = False
If PF(9).Visible = True Then PF(9).Visible = False
If PF(15).Visible = True Then PF(15).Visible = False
Call CMV(Me)
DoEvents
lRet = SetInitEntry("MsgBOX", "LEFT", Me.Left - 100)
lRet = SetInitEntry("MsgBOX", "TOP", Me.Top + (Me.Height - 2400) / 2)
End Sub

Private Sub LBSINGER_Click()
If LBSINGER.Caption = "" Then Exit Sub
Clipboard.Clear
Clipboard.SetText LBSINGER.Caption
End Sub

Private Sub LBSINGER_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call UpNow
End Sub

Private Sub LBSINGER_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LBSINGER.ToolTipText = LBSINGER.Caption
End Sub

Private Sub LBSONG_Change()
LA(18).Caption = LBSONG.Caption
End Sub

Private Sub LBSONG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call UpNow
End Sub

Private Sub LBSONG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Wm.playState = wmppsPlaying Then LBSONG.ToolTipText = SONGNAME
End Sub

Private Sub LBSONG_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub
Private Sub lbthing_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
If Frmm.LSTLINK.List(0) = "" Then
Frmm.WB.Navigate "http://hi.baidu.com/iceeorgan/item/96d45007a86c1acbff240dfa"
Frmm.WB.Refresh
Call UpNow
Exit Sub
Else
ShellExecute 0&, vbNullString, Split(Frmm.LSTLINK.List(0), "|")(1), vbNullString, vbNullString, 0 '调用ie
End If
End Sub

Private Sub lbthing_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PF(15).Visible = True Then PF(15).Visible = False
End Sub

Private Sub lbthing_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub
Private Sub lbuse_Change()
Call SaveSetting("ICEE", "MAIN", "USERNAME", LBUSE.Caption)
End Sub

Private Sub lbuse_DblClick()
If UNME.Enabled = True Then Call NoNoNo
End Sub

Private Sub LBUSE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Call UpNow
End Sub
Private Sub lbuse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LBUSE.ToolTipText = LBUSE.Caption
End Sub
Private Sub lbzt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Winsock1.State <> 7 Then Exit Sub
If HAS_HEAD = False Then Call CMV(Me): Exit Sub
If PF(15).Visible = True Then PF(15).Visible = False
If Button = 1 Then Me.PopupMenu Frmm.iM, , lbzt.Left - 100, lbzt.Top + lbzt.Height + 5
End Sub

Private Sub lbzt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub lbzt_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Tit_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub LCO_Change()
lRet = SetInitEntry("IM", "MSGCOUNT", LCO.Caption)
LA(30).Caption = LCO.Caption
End Sub

Private Sub LCO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMGEM.PICTURE = Frmm.PIC(68).PICTURE Then IMGEM.PICTURE = Frmm.PIC(17).PICTURE
If IMSKIN.PICTURE <> Frmm.PIC(12).PICTURE Then IMSKIN.PICTURE = Frmm.PIC(12).PICTURE
End Sub

Private Sub LCO_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If frmma.Left > FRMHIS.Width Then
FRMHIS.Move frmma.Left - FRMHIS.Width, frmma.Top
Else
FRMHIS.Move frmma.Left + frmma.Width, frmma.Top
End If
FRMHIS.Show
GETMSGCOUNT = 0
LCO.Caption = 0
End Sub

Private Sub LCO_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Tit_OLEDragDrop(Data, Effect, Button, Shift, X, Y)

End Sub

Private Sub ld1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ld1.Visible = False
ld2.Visible = True
End Sub
Private Sub ld2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
ld2.Visible = False
ld3.Visible = True
End If
End Sub
Private Sub ld3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ld3.Visible = False
ld1.Visible = True
If ld3.Visible = False Then
PF(3).ZOrder 0
PF(3).Move 0, 0, iFrame.ScaleWidth, iFrame.ScaleHeight
pl.Cls
pl.Move 0, 0, PF(3).ScaleWidth, PF(3).ScaleHeight
ICZ(3).Visible = True
IS_PICSHOW = True
End If
End Sub
Private Sub LISTBAIDU_Click()
TXTBAIDU.Text = LISTBAIDU.List(LISTBAIDU.ListIndex)
ShellExecute 0&, vbNullString, URLTMP & TXTBAIDU.Text, vbNullString, vbNullString, 0   '调用ie
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With ListView1
   If (ColumnHeader.Index - 1) = .SortKey Then
  .SortOrder = (.SortOrder + 1) Mod 2
  .Sorted = True
   Else
  .Sorted = False
  .SortOrder = 0
  .SortKey = ColumnHeader.Index - 1
  .Sorted = True
   End If
End With
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Call SHOWDEMO
End Sub
Sub SHOWDEMO()
On Error Resume Next
If ListView1.SelectedItem.Index < 1 Then Exit Sub
Dim C As clsClip
Set C = CLIPS.Item(ListView1.SelectedItem.Index)
If Not C Is Nothing Then
If C.IsImage Then
txtText.Visible = False
ICT.Visible = False
Set ImgPreview.PICTURE = C.image
If ImgPreview.Height > PVP.ScaleHeight Or ImgPreview.Width > PVP.ScaleWidth Then
IPR.Height = PVP.ScaleHeight
IPR.Width = PVP.ScaleWidth * (IPR.Height / ImgPreview.ScaleHeight)
Dimention2 IPR, ImgPreview, ImgPreview.ScaleWidth * (IPR.Height / ImgPreview.ScaleHeight), IPR.Height
IPR.Move (PVP.ScaleWidth - IPR.Width) / 2, 0
Else
Dimention2 IPR, ImgPreview, ImgPreview.Width, ImgPreview.Height
IPR.Move (PVP.ScaleWidth - IPR.Width) / 2, (PVP.ScaleHeight - IPR.Height) / 2
End If
PVP.ZOrder 0
Else
ICT.Visible = True
ICT.ZOrder 0
txtText.ZOrder 0
txtText.Visible = True
txtText.Text = C.ClipText
End If
End If
Set C = Nothing
End Sub
Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub
Private Sub lstRes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lstRes.ListCount = 0 Then Exit Sub
If Button = 2 Then Me.PopupMenu Frmm.主机
End Sub

Private Sub lstRes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub
Private Sub Mbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpNow
End Sub

Private Sub Mbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub
Sub LEAVEPIC()
PICD.Visible = False
IMJ.Visible = False
If AUTOPLAYPIC = True Then
Timers.Enabled = True
End If
PF(6).Visible = True
pl.AutoRedraw = False
LOCKSAFE
End Sub
Sub ShowMusic()
BACKME.Visible = True
PP.Visible = True
Timers.Enabled = False
PP.ZOrder 0
PICDL.ZOrder 0
PICD.Visible = False
PF(6).Visible = False
Pmusic.Visible = False
PICSER.Visible = False
Call DRAWMUSIC
End Sub

Private Sub MBK_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
On Error Resume Next
IMSIGN.Visible = True
USE_PIC_FORM = True
lRet = SetInitEntry("SYSTEM", "BACKPICTURE", App.Path + "\SKIN\BK\" & Index & ".JPG")
lRet = SetInitEntry("SYSTEM", "BACKPICTURE_INDEX", Index)
Frmm.IMBK.PICTURE = LoadPicture(App.Path + "\SKIN\BK\" & Index & ".JPG")
Call SUBDRAW
lRet = SetInitEntry("SYSTEM", "USE_PIC", USE_PIC_FORM)
End Sub

Private Sub MINIME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If UNME.PICTURE <> Frmm.PIC(177).PICTURE Then UNME.PICTURE = Frmm.PIC(177).PICTURE
If BACKME.PICTURE <> Frmm.PIC(175).PICTURE Then BACKME.PICTURE = Frmm.PIC(175).PICTURE
If MINIME.PICTURE <> Frmm.PIC(180).PICTURE Then MINIME.PICTURE = Frmm.PIC(180).PICTURE
If SETME.PICTURE <> Frmm.PIC(173).PICTURE Then SETME.PICTURE = Frmm.PIC(173).PICTURE

End Sub

Private Sub MINIME_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Call NoNoNo
End Sub


Private Sub MNO_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
USE_PIC_FORM = False
IMSIGN.Visible = False
Set iFrame.PICTURE = Nothing
Call DRAWFACE
Call SUBDRAW
lRet = SetInitEntry("SYSTEM", "USE_PIC", USE_PIC_FORM)
End Sub

Private Sub PBK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CTL_MOVE = False Then
CTL_MOVE = True
PBK.Cls
Call PaintPng(App.Path & "\SKIN\BK_H.PNG", PBK.hdc, 0, 0)
End If
End Sub

Private Sub PBK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
PF(11).Visible = False
Call LOCKSAFE
End Sub

Private Sub pc_KeyDown(KeyCode As Integer, Shift As Integer)
Call IFRAME_KeyDown(KeyCode, 0)
End Sub

Private Sub PC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Call UpNow
End If
End Sub

Private Sub Pc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PF(15).Visible = True Then PF(15).Visible = False
MOVENOW
End Sub

Private Sub pc_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub PCOO_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If D_L_SHOW = False Then Exit Sub
Call FrmNetMusic.SETLRCCOLOR(Index + 1)
FrmNetMusic.cDeskLrc.ReDraw
lRet = SetInitEntry("PLAYER", "LRCSHOW_COLOR", Index + 1)
End Sub

Private Sub PDB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpNow
End Sub

Private Sub PDB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MOVENOW
End Sub
Private Sub pe1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
pe1.Visible = False
pe2.Visible = True
End Sub

Private Sub pe2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
pe2.Visible = False
pe3.Visible = True
End If
End Sub
Private Sub pe3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
pe3.Visible = False
pe1.Visible = True
If pe3.Visible = False Then
Call RUNSAFE
PICD.Cls
LES = BitBlt(PICD.hdc, 0, 0, PICD.Width, PICD.Height, iFrame.hdc, PICD.Left, PICD.Top, &HCC0020)
PICD.Line (0, 0)-(PICD.ScaleWidth, 64), Frmm.PTCO.POINT(1, 1), BF
Call PaintPng(App.Path & "\SKIN\PHOTOSET.png", PICD.hdc, 0, 0) '幻灯片设置
IMJ.BackColor = Frmm.PTCO.POINT(1, 1)
ICL(5).SETCOLOR Frmm.PTCO.POINT(1, 1), &H554E4, vbWhite
AUTOPLAYPIC = GetInitEntry("SYSTEM", "AUTOPLAYPIC", True)
If AUTOPLAYPIC = True Then IMR.PICTURE = Frmm.PIC(130).image Else IMR.PICTURE = Frmm.PIC(131).image
PF(6).Visible = False
PICD.Visible = True
PICD.ZOrder 0
IMJ.ZOrder 0
PICTOOL.ZOrder 0
Timers.Enabled = False
IMJ.Visible = True
PICD.Refresh
If AUTOPLAYPIC = True Then

Else

End If
End If
End Sub

Private Sub PF_DblClick(Index As Integer)
If Index <> 4 Then Exit Sub
If PF(17).Visible = False Then PF(17).Visible = True
End Sub

Private Sub PF_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Call IFRAME_KeyDown(KeyCode, 0)

End Sub

Private Sub PF_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Call UpNow
End Sub

Private Sub PF_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim I As Integer
Call REDRAW_PLAY_CON
If IMJ.PICTURE <> Frmm.X1.PICTURE Then IMJ.PICTURE = Frmm.X1.PICTURE
If MUSIC_MOVE = True Then
MUSIC_MOVE = False
PICBACK.Cls
PICDL.Cls
IU.Cls
If PICBACK.BackColor <> PP.BackColor Then PICBACK.BackColor = PP.BackColor
If IU.BackColor <> PP.BackColor Then IU.BackColor = PP.BackColor
If PICDL.BackColor <> COLOR_NOR Then PICDL.BackColor = COLOR_NOR
LES = BitBlt(PICBACK.hdc, 0, 0, PICBACK.Width, PICBACK.Height, PP.hdc, PICBACK.Left, PICBACK.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\L_N.PNG", PICBACK.hdc, 8, 6)
PICBACK.Refresh
LES = BitBlt(PICDL.hdc, 0, 0, PICDL.Width, PICDL.Height, iFrame.hdc, PICDL.Left, PICDL.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\D_N.PNG", PICDL.hdc, 5, 3)
PICDL.Refresh
LES = BitBlt(IU.hdc, 0, 0, IU.Width, IU.Height, Pmusic.hdc, IU.Left, IU.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\MB_N.PNG", IU.hdc, 5, 3)
IU.Refresh
End If
pe1.Visible = True
pe2.Visible = False
pe3.Visible = False
ld1.Visible = True
ld2.Visible = False
ld3.Visible = False
End Sub
Private Sub PF_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub
Private Sub PICAD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpNow
End Sub

Private Sub PICAD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub
Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MUSIC_MOVE = False Then
MUSIC_MOVE = True
PICBACK.Cls
If PICBACK.BackColor <> PP.BackColor Then PICBACK.BackColor = PP.BackColor
LES = BitBlt(PICBACK.hdc, 0, 0, PICBACK.Width, PICBACK.Height, PP.hdc, PICBACK.Left, PICBACK.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\L_H.PNG", PICBACK.hdc, 8, 6)
PICBACK.Refresh
End If
End Sub

Private Sub PICBACK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Call SHOWUIT
End Sub
Sub SHOWUIT()
Pmusic.ZOrder 0
Pmusic.Visible = True
lRet = SetInitEntry("PLAYER", "MUSICMODE", 0)
End Sub
Private Sub PICBACK_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub
Private Sub PICBUG_KeyDown(KeyCode As Integer, Shift As Integer)
Call IFRAME_KeyDown(KeyCode, 0)
End Sub
Private Sub PICBUG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpNow
End Sub

Private Sub PICBUG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW

End Sub



Private Sub PICBUG_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y): Call DRAWBUG
End Sub
Sub 物理删除歌曲()
On Error Resume Next
Dim OP As SHFILEOPSTRUCT
If Wm.URL = WILL_DEL Then Exit Sub
PLIST.RemoveItem (WILL_DEL_IDX)
With OP
.wFunc = FO_DELETE
.pFrom = WILL_DEL
.fFlags = FOF_ALLOWUNDO
End With
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">物理删除了" & WILL_DEL
SHFileOperation OP
Call SAVELIST
End Sub
Sub 去除重复()
On Error Resume Next
Dim n As Integer, m As Integer
For n = 0 To PLIST.ListCount - 1
For m = n To PLIST.ListCount - 1
If PLIST.URL(n) = PLIST.URL(m) And m <> n Then PLIST.RemoveItem m
Next
Next
Call SAVELIST
End Sub
Private Sub PICCLIP_KeyDown(KeyCode As Integer, Shift As Integer)
Call IFRAME_KeyDown(KeyCode, 0)

End Sub

Private Sub PICCLIP_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y): Call DRAWCLIP
End Sub

Private Sub Picd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Call LEAVEPIC
End Sub
Private Sub Picd_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub
Private Sub PICDL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMSERG.Visible = True Then IMSERG.Visible = False
If MUSIC_MOVE = False Then
MUSIC_MOVE = True
PICDL.Cls

If PICDL.BackColor <> COLOR_NOR Then PICDL.BackColor = COLOR_NOR
LES = BitBlt(PICDL.hdc, 0, 0, PICDL.Width, PICDL.Height, iFrame.hdc, PICDL.Left, PICDL.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\D_H.PNG", PICDL.hdc, 5, 3)
PICDL.Refresh
End If
End Sub

Private Sub PICDL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Call SHOWDL
End Sub

Private Sub PICDL_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub
Private Sub PICFI_KeyDown(KeyCode As Integer, Shift As Integer)
Call IFRAME_KeyDown(KeyCode, 0)

End Sub

Private Sub PICFI_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y): Call DrawInfo
End Sub


Private Sub PICIG_KeyDown(KeyCode As Integer, Shift As Integer)
Call IFRAME_KeyDown(KeyCode, 0)

End Sub

Private Sub PICIG_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y): Call DRAWUN
End Sub

Private Sub PICIM_KeyDown(KeyCode As Integer, Shift As Integer)
Call IFRAME_KeyDown(KeyCode, 0)

End Sub

Private Sub PICIM_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub
Private Sub PICLO_KeyDown(KeyCode As Integer, Shift As Integer)
Call IFRAME_KeyDown(KeyCode, 0)
End Sub

Private Sub PICLO_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub PICMU_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub PICNET_KeyDown(KeyCode As Integer, Shift As Integer)
Call IFRAME_KeyDown(KeyCode, 0)

End Sub

Private Sub PICNET_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpNow
End Sub

Private Sub PICNET_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub PICNET_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y): Call DRAWNET
End Sub

Private Sub PICPASS_KeyDown(KeyCode As Integer, Shift As Integer)
Call IFRAME_KeyDown(KeyCode, 0)

End Sub

Private Sub PICPASS_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y): Call DRAWPASS
End Sub

Private Sub PICSER_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpNow

End Sub

Private Sub PICSER_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub PICTIME_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
'If TMP.Enabled = True Then
'If UNCOUNT = False Then
'UNCOUNT = True
'Else
'UNCOUNT = False
'End If
'Else
UpNow
'End If
End Sub

Private Sub PICTIME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub Picture1_Resize()
Picture2.Cls
Picture2.Height = Picture1.Height
Picture2.Width = Picture1.Width
PLOGO.PaintPicture Picture1.image, 0, 0, PLOGO.Width, PLOGO.Height
Call PaintPng(App.Path & "\SKIN\HEAD.png", PLOGO.hdc, 0, -1)
DrawGray
Call LoadParam
End Sub
Private Sub picClip_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpNow
End Sub

Private Sub picClip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub
Private Sub PICFI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpNow
End Sub

Private Sub PICFI_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub PICIG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpNow
End Sub

Private Sub PICIG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub PICIM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y >= 0 And Y <= 30 Then
UpNow
Else
PF(3).ScaleMode = 1
gX = X
End If
End Sub

Private Sub PICIM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW

If Not Button = vbLeftButton Then Exit Sub
If SETME.Enabled = True And PF(3).ScaleMode = 1 Then

Dim dX As Long, dY As Long, ax As Long, ay As Long, t As Long, L As Long, tt As Long, ll As Long
     With PICIM
  dX = X - gX
  ll = .Left
  L = Abs(ll)
  ax = (.Width - L) '- ScaleWidth)
  If ll > 0 Then
  Else
       If dX < 100 Then
       Else
   If dX > L Then dX = L
       End If
  End If
  .Move Abs(ll + dX)
 
  pl.Left = PICIM.Left - pl.Width
      End With
End If
End Sub

Private Sub PICIM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If PICIM.Left > 0 And PICIM.Left <= 199 Then PICIM.Left = 0: pl.Left = PICIM.Left - pl.Width
If PICIM.Left > 200 Then Call ShowPic

End Sub
Private Sub PicLo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y > 0 And Y < 30 Then
UpNow
Else
PF(3).ScaleMode = 1
gX = X
End If
End Sub

Private Sub PicLo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
If Not Button = vbLeftButton Then Exit Sub
If PF(3).ScaleMode = 1 And SETME.Enabled = True Then
Dim dX As Long, dY As Long, ax As Long, ay As Long, t As Long, L As Long, tt As Long, ll As Long
     With PICIM
  dX = X - gX
  ll = .Left
  L = Abs(ll)
  ax = (.Width - L) '- ScaleWidth)
  If ll > 0 Then
  Else
       If dX < 100 Then
       Else
   If dX > L Then dX = L
       End If
  End If
  .Move Abs(ll + dX)
 
  pl.Left = PICIM.Left - pl.Width
      
      End With
End If
End Sub

Private Sub PicLo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If PICIM.Left > 0 And PICIM.Left <= 199 Then PICIM.Left = 0: pl.Left = PICIM.Left - pl.Width
If PICIM.Left > 200 Then Call ShowPic

End Sub
Private Sub PICMU_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PF(15).Visible = True Then PF(15).Visible = False
ld1.Visible = True
ld2.Visible = False
ld3.Visible = False
PICMU.Visible = False
PNZ.Visible = True
If TMP.Enabled = True And USE_PIC_FORM = True Then PICMU.Visible = False: PNZ.Visible = True: Exit Sub
Do While MOUSEMO = False
MOUSEMO = True
Frmm.TMMU.Enabled = True
PICMU.PICTURE = LoadPicture("")
ITIME = 0 'PICMU.Picture = Frmm.da2.Picture
Loop
End Sub
Private Sub PICPASS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpNow
End Sub

Private Sub PICPASS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub
Private Sub PICTOOL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpNow
End Sub

Private Sub PICTOOL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PF(15).Visible = True Then PF(15).Visible = False
MOVENOW
End Sub


Private Sub PicUse_KeyDown(KeyCode As Integer, Shift As Integer)
Call IFRAME_KeyDown(KeyCode, 0)

End Sub

Private Sub PICUSE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y >= 0 And Y <= 30 Then
Call UpNow
Else
PF(3).ScaleMode = 1
gX = X
End If
End Sub

Private Sub PicUse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
If Not Button = vbLeftButton Then Exit Sub
If SETME.Enabled = True = True And PF(3).ScaleMode = 1 Then
     Dim dX As Long, dY As Long, ax As Long, ay As Long, t As Long, L As Long, tt As Long, ll As Long
     With PicUse
  dX = X - gX
  ll = .Left
  L = Abs(ll)
  ax = (.Width - L - ScaleWidth)
  If ll > 0 Then
       dX = 0
  Else
       If dX < 0 Then
       Else
    If dX > L Then dX = L
       End If
  End If
  .Move ll + dX
  pl.Move PicUse.Left + PicUse.Width
     End With
End If
End Sub

Private Sub PicUse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PicUse.Left > -200 And PicUse.Left < 0 Then PicUse.Left = 0: pl.Left = PicUse.Left + PicUse.Width
If PicUse.Left <= -199 Then Call ShowPic
PF(3).ScaleMode = 3
End Sub

Private Sub PicUse_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim strpath As String, Str1 As String
If Data.files.Count > 0 Then strpath = Data.files(1)
Select Case LCase$(Right$(Data.files(1), 3))
Case "bmp", "jpg", "psd", "png", "gif"
Call FRMBOARD.OpenFile(strpath)
Case "mp3"
PLIST.AddItem LastFileName(strpath), "", strpath, 0
Call SAVELIST
Wm.URL = strpath
Case "m3u"
Call Playlist(strpath)
Case "txt", "lrc", "sri"
Dim nW As New FrmNew
nW.Show
nW.TXTTS.LoadFile (strpath)
Case Else
Call OpenFile(strpath)
End Select
End Sub

Private Sub PicZoom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call UpNow
End Sub

Private Sub PicZoom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW

End Sub

Private Sub PKU_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CTL_MOVE = False Then
CTL_MOVE = True
LES = BitBlt(PKU.hdc, 0, 0, PKU.Width, PKU.Height, PP.hdc, PKU.Left, PKU.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\KU_H.PNG", PKU.hdc, 0, 0)
PKU.Refresh
End If
End Sub

Private Sub PKU_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Call MUSICBOX
End Sub

Private Sub pl_KeyDown(KeyCode As Integer, Shift As Integer)
Call IFRAME_KeyDown(KeyCode, 0)

End Sub

Private Sub pl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pl.Left >= 800 Then Call ShowFriend
If pl.Left > 0 And pl.Left <= 799 Then pl.Left = 0: PicUse.Left = pl.Left - PicUse.Width: PICIM.Left = pl.Left + pl.Width
If pl.Left < 0 And pl.Left >= -798 Then pl.Left = 0: PicUse.Left = pl.Left - PicUse.Width: PICIM.Left = pl.Left + pl.Width
If pl.Left <= -799 Then Call ShowIM
PF(3).ScaleMode = 3
End Sub
Private Sub pL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IS_PICSHOW = True Then Call UpNow: Exit Sub
If Y >= 0 And Y <= 30 Then
UpNow
Else
PF(3).ScaleMode = 1
If Button = 1 Then gX = X
End If
End Sub
Private Sub pL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MOVENOW
If IS_PICSHOW = True And Y > 0 And Y <= 50 Then ICZ(3).Visible = True Else ICZ(3).Visible = False
If IS_PICSHOW = True Then Exit Sub
If Not Button = vbLeftButton Then Exit Sub
     Dim dX As Long, dY As Long, ax As Long, ay As Long, t As Long, L As Long, tt As Long, ll As Long
     With pl
  dX = X - gX
  ll = .Left
  L = Abs(ll)
  ax = (.Width - L - ScaleWidth)
  If ll >= 800 Then
       dX = 0
  Else
       If dX < 0 Then
    If Abs(dX) > ax Then dX = -ax
       Else
    'If dX > L Then dX = L
       End If
  End If
  .Move ll + dX
  PicUse.Move pl.Left - PicUse.Width
  PICIM.Left = pl.Left + pl.Width
     End With
End Sub


Private Sub picd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpNow
End Sub

Private Sub Picd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MOVENOW
End Sub
Private Sub pl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y): DRAWMUSIC
End Sub
Private Sub PLAYB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Pmusic.Visible = True Then Pmusic.Visible = False
If PV.Visible = True Then PV.Visible = False
If CTL_MOVE = False Then
CTL_MOVE = True
PLAYB.Cls
IPRE.Cls
INEXT.Cls
INEXT.BackColor = PP.BackColor
IPRE.BackColor = PP.BackColor
PLAYB.BackColor = PP.BackColor
LES = BitBlt(PLAYB.hdc, 0, 0, PLAYB.Width, PLAYB.Height, PP.hdc, PLAYB.Left, PLAYB.Top, &HCC0020)
If Wm.playState = wmppsPlaying Then Call PaintPng(App.Path & "\SKIN\PA_H.PNG", PLAYB.hdc, 0, 0) Else Call PaintPng(App.Path & "\SKIN\P_H.PNG", PLAYB.hdc, 0, 0)
PLAYB.Refresh
LES = BitBlt(IPRE.hdc, 0, 0, IPRE.Width, IPRE.Height, PP.hdc, IPRE.Left, IPRE.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\PR_N.PNG", IPRE.hdc, 0, 0)
IPRE.Refresh
LES = BitBlt(INEXT.hdc, 0, 0, INEXT.Width, INEXT.Height, PP.hdc, INEXT.Left, INEXT.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\NX_N.PNG", INEXT.hdc, 0, 0)
INEXT.Refresh
End If
End Sub

Private Sub PLAYB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
If Wm.playState = wmppsPaused Or Wm.playState = wmppsStopped Then
PLAYB.Cls
LES = BitBlt(PLAYB.hdc, 0, 0, PLAYB.Width, PLAYB.Height, PP.hdc, PLAYB.Left, PLAYB.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\PA_N.PNG", PLAYB.hdc, 0, 0)
PLAYB.Refresh
TMP.Enabled = True
If Wm.URL = "" Then Wm.URL = GetInitEntry("PlayList", "LastUrl", "")
Wm.Controls.Play
Else
PLAYB.Cls
LES = BitBlt(PLAYB.hdc, 0, 0, PLAYB.Width, PLAYB.Height, PP.hdc, PLAYB.Left, PLAYB.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\P_N.PNG", PLAYB.hdc, 0, 0)
PLAYB.Refresh
Wm.Controls.pause
TMP.Enabled = False
PICMU.Cls
PICMU.PICTURE = Frmm.da1.PICTURE
End If
End Sub

Private Sub PLAYB_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub PMDL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CTL_MOVE = False Then
CTL_MOVE = True
LES = BitBlt(PMDL.hdc, 0, 0, PMDL.Width, PMDL.Height, PP.hdc, PMDL.Left, PMDL.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\DD_H.PNG", PMDL.hdc, 0, 0)
PMDL.Refresh
End If
End Sub

Private Sub PMDL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
On Error GoTo ERR
Call DoFileDownload(StrConv(Wm.URL, vbUnicode))
ERR:
Exit Sub
End Sub

Private Sub PMINFO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CTL_MOVE = False Then
CTL_MOVE = True
LES = BitBlt(PMINFO.hdc, 0, 0, PMINFO.Width, PMINFO.Height, PP.hdc, PMINFO.Left, PMINFO.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\MI_H.PNG", PMINFO.hdc, 0, 0)
PMINFO.Refresh
End If
End Sub

Private Sub PMINFO_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If Wm.URL = "" Then Exit Sub
Call FRMMIN.SeeIt(Wm.URL)
If Me.Left > FRMMIN.Width Then
FRMMIN.Move Me.Left - FRMMIN.Width, Me.Top
Else
FRMMIN.Move Me.Left + Me.Width, Me.Top
End If
FRMMIN.Show
End Sub

Private Sub PMU_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Select Case Index
Case 0
If frmma.Left > FRMFAV.Width Then
FRMFAV.Move frmma.Left - FRMFAV.Width, frmma.Top
Else
FRMFAV.Move frmma.Left + frmma.Width, frmma.Top
End If
FRMFAV.Show
Case 1
FrmNetMusic.Show
Case 3
FrmNetMusic.PO(7).Visible = True
FrmNetMusic.PO(7).ZOrder 0
FrmNetMusic.Show
End Select
End Sub

Private Sub Pmusic_Resize()
Mbar.Move 0, Pmusic.ScaleHeight - Mbar.Height
End Sub

Private Sub PTIME_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub
Private Sub PLIST_Click(Button As Integer, Shift As Integer, X As Single, Y As Single)
WILL_DEL = PLIST.URL(PLIST.ListIndex)
WILL_DEL_IDX = PLIST.ListIndex
If Button = 2 Then PLIST.SetFocus: Me.PopupMenu Frmm.文件
End Sub

Private Sub PLIST_DBClick(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call 播放歌曲
End Sub

Private Sub PLIST_PlayIndexChanged(OldIndex As Integer, NewIndex As Integer, wFlag As Integer)
Song = PLIST.PlayIndex  '将播放歌曲位置储存
If IS_NET = True Then Call FrmNetMusic.RELIST
lRet = SetInitEntry("PlayList", "LastUrl", Wm.URL)   '记录退出时音乐播放器的路径
lRet = SetInitEntry("Playlist", "LastIndex", PLIST.PlayIndex)
End Sub

Private Sub PLIST_Resize()
PLIST.ListIndex = 0
PLIST.Refresh
End Sub

Private Sub Pmusic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpNow
End Sub

Private Sub Pmusic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MOVENOW
End Sub

Private Sub Pmusic_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y): Call DRAWMUSIC
End Sub

Private Sub PNZ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PNZ.PICTURE <> Frmm.da2.image Then PNZ.PICTURE = Frmm.da2.image
End Sub

Private Sub PNZ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PF(15).Visible = True Then PF(15).Visible = False
If PNZ.PICTURE <> Frmm.da2.image Then PNZ.PICTURE = Frmm.da2.image
ld1.Visible = True
ld2.Visible = False
ld3.Visible = False
End Sub

Private Sub PNZ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
PNZ.Visible = False
PICMU.Visible = True
IS_M_S = True
frmmp.Left = Me.Left + 100
frmmp.Top = Me.Top + 3760
frmmp.Show
End If
End Sub

Private Sub PNZ_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Sub 播放歌曲()
If PLIST.ListCount = 0 Then Exit Sub
Wm.URL = PLIST.URL(PLIST.ListIndex) '加载歌曲
Song = PLIST.ListIndex '将播放歌曲位置储存
Wm.Controls.Play '播放
End Sub
Private Sub PLIST_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MOVENOW
End Sub
Private Sub PP_KeyDown(KeyCode As Integer, Shift As Integer)
Call IFRAME_KeyDown(KeyCode, 0)
End Sub
Private Sub PP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpNow
End Sub
Private Sub pp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
If FILESINGER = "" Then Exit Sub
 If PathFileExists(App.Path & "\MEDIA\MusicPicture\" & FILESINGER & ".Bmp") = 1 Then Exit Sub
Call Frmm.CHECKNET
If Status.RasConnState <> &H2000 Then Exit Sub
If IMSERG.Visible = False Then IMSERG.Visible = True
If IMCLEAR.Visible = True Then PMINFO.Left = IMCLEAR.Width + IMCLEAR.Left + 5 Else PMINFO.Left = IMCLEAR.Left
If Pmusic.Visible = True Then Pmusic.Visible = False

End Sub
Private Sub PP_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim strpath As String
If Data.files.Count > 0 Then
strpath = Data.files(1)
Select Case LCase(Right(Data.files(1), 3))
Case "bmp", "gif", "jpg"
If Wm.URL = "" Then Exit Sub
If FILESINGER = "" Then Exit Sub
FileCopy strpath, App.Path & "\MEDIA\MusicPicture\" & FILESINGER & ".Bmp"
Call GETSINGER
Case "png"
If Wm.URL = "" Then Exit Sub
If FILESINGER = "" Then Exit Sub
Call OPENISPNG(Frmm.PSINGER, strpath)
Call PictureBoxSaveJPG(Frmm.PSINGER.image, App.Path & "\MEDIA\MusicPicture\" & FILESINGER & ".Bmp", 100)  '保存歌手图像
Call GETSINGER
Case "mp3"
PLIST.AddItem LastFileName(strpath), "", strpath, 0
Call SAVELIST
Wm.URL = strpath
Case "m3u"
Call Playlist(strpath)
'Case "rar", "zip", "7z", "apk"
Case Else
Call OpenFile(strpath)
End Select
End If
End Sub
Private Sub PSEND_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CTL_MOVE = False Then
CTL_MOVE = True
LES = BitBlt(PSEND.hdc, 0, 0, PSEND.Width, PSEND.Height, PP.hdc, PSEND.Left, PSEND.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\SEND_H.PNG", PSEND.hdc, 0, 0)
PSEND.Refresh
End If
End Sub

Private Sub PSEND_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
DoEvents
On Error GoTo ERR
If GetUdisk() = "" Then Exit Sub
Call FileCopy(Wm.URL, GetUdisk() & ":\" & LastFileName(Wm.URL))
ERR:
Exit Sub
End Sub

Private Sub Pser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpNow
End Sub

Private Sub Pser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub
Private Sub OnMouseGestureStart(ByVal XPos As Long, ByVal YPos As Long)
    m_xPos = XPos
    m_yPos = YPos
    Set m_GeCol = Nothing       '清空上次的轨迹点集合
    Set m_GeCol = New Collection
End Sub

Private Sub OnMouseGestureEnd(ByVal XPos As Long, ByVal YPos As Long)
    '
End Sub
Private Sub OnMouseGesture(ByVal XPos As Long, ByVal YPos As Long)
On Error Resume Next
    Dim xVal    As Long '坐标差值
    Dim yVal    As Long
    Dim sTemp   As String       '点的XY字符串标记，0表示该点没有移动，L左，R右，U上，D下
    
    xVal = XPos - m_xPos   '计算轨迹的差值
    yVal = YPos - m_yPos
    
    If Abs(xVal) >= 50 Or Abs(yVal) >= 20 Then  '当移动幅度超过10个像素，记录下这个点
'以x坐标差值为参考记录该点
If Abs(xVal) <= 10 Then  'x坐标幅度在3像素以内认为没有移动
    sTemp = IIf(yVal > 0, "0D", "0U")
ElseIf xVal > 0 Then    'x坐标差值小于0表示往左移动
    If Abs(yVal) <= 10 Then  'y坐标差值小于3认为没有移动
sTemp = "R0"
    ElseIf yVal > 15 Then
sTemp = "RD"
    Else
sTemp = "RU"
    End If
ElseIf xVal < 0 Then    'x坐标差值大于0表示往右移动
    If Abs(yVal) <= 30 Then
sTemp = "L0"
    ElseIf yVal > 50 Then
sTemp = "LD"
    Else
sTemp = "LU"
    End If
End If

m_GeCol.Add sTemp
m_xPos = XPos
m_yPos = YPos

Dim sGesture As String
sGesture = GetMouseGestureString

If InStr(1, sGesture, "R0R0R0") Then
  NT (2) '"向右"
ElseIf InStr(1, sGesture, "L0L0L0") Then
  NT (1) '"向左"
ElseIf InStr(1, sGesture, "0U0U0U") Then
    '"向上"
ElseIf InStr(1, sGesture, "0D0D0D") Then
  '"向下"
ElseIf sGesture Like "R*U*L*D" Then
   ' "逆时针圈"
ElseIf sGesture Like "LDLD*RD" Then
   ' "交叉"
End If
    End If
End Sub
Private Function GetMouseGestureString() As String
On Error Resume Next
    '获取手势字符串
    Dim nCount      As Long
    Dim sGesture    As String
    Dim nPos As Long
    nCount = m_GeCol.Count
    nPos = Fix(nCount / 4) + 1  '取点间隔的个数
    If nCount > 0 Then  '取4个点
sGesture = m_GeCol.Item(1)
sGesture = sGesture & m_GeCol.Item(nPos)
sGesture = sGesture & m_GeCol.Item(nPos * 2)
sGesture = sGesture & m_GeCol.Item(nCount)
    End If
    GetMouseGestureString = sGesture
End Function
Private Sub PSubClass_MsgCome(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, lng_hWnd As Long, uMsg As Long, wParam As Long, lParam As Long)
If bBefore Then
If uMsg = WM_MOVE Then MoveWindow frmmabk.hwnd, Me.Left / 15 - 20, Me.Top / 15 - 10, 380, 650, True
End If
End Sub

Private Sub PV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then ISMOVE = True
End Sub

Private Sub PV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERR
If ISMOVE = True Then
ML.Width = X
If ML.Width <= 0 Then ML.Width = 0: Exit Sub
If ML.Width >= PV.ScaleWidth Then ML.Width = PV.ScaleWidth: ISMOVE = False
Wm.settings.volume = Int((100 / PV.ScaleWidth) * ML.Width)
End If
PV.ToolTipText = "音量:" & Wm.settings.volume & "%"
ERR:

End Sub

Private Sub PV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ISMOVE = False
lRet = SetInitEntry("PLAYER", "VOLUME", ML.Width)
End Sub

Private Sub PVP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpNow
End Sub

Private Sub PVP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Integer
For I = 0 To ICC.Count - 1
If ICC(I).Visible = False Then ICC(I).Visible = True
Next
End Sub

Private Sub PZOR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CTL_MOVE = False Then
CTL_MOVE = True
PZOR.Cls
PZOR.BackColor = PP.BackColor
LES = BitBlt(PZOR.hdc, 0, 0, PZOR.Width, PZOR.Height, PP.hdc, PZOR.Left, PZOR.Top, &HCC0020)
Select Case LOLIPOP
Case 3
Call PaintPng(App.Path & "\SKIN\SX_H.PNG", PZOR.hdc, 0, 0)
Case 1
Call PaintPng(App.Path & "\SKIN\DQ_H.PNG", PZOR.hdc, 0, 0)
Case 2
Call PaintPng(App.Path & "\SKIN\XH_H.PNG", PZOR.hdc, 0, 0)
Case 0
Call PaintPng(App.Path & "\SKIN\SJ_H.PNG", PZOR.hdc, 0, 0)
End Select
PZOR.Refresh
End If
End Sub

Private Sub PZOR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Me.PopupMenu Frmm.顺序, , PZOR.Left, PZOR.Top + 160
End Sub

Private Sub SETME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If UNME.PICTURE <> Frmm.PIC(177).PICTURE Then UNME.PICTURE = Frmm.PIC(177).PICTURE
If BACKME.PICTURE <> Frmm.PIC(175).PICTURE Then BACKME.PICTURE = Frmm.PIC(175).PICTURE
If MINIME.PICTURE <> Frmm.PIC(179).PICTURE Then MINIME.PICTURE = Frmm.PIC(179).PICTURE
If SETME.PICTURE <> Frmm.PIC(174).PICTURE Then SETME.PICTURE = Frmm.PIC(174).PICTURE
End Sub

Private Sub SETME_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
frmset.Show
End Sub

Private Sub SHRO_Change()
PF(4).Top = -SHRO.Value
End Sub

Private Sub SHRO_Scroll()
SHRO_Change
End Sub

Private Sub SURO_Change()
PF(16).Top = -SURO.Value + 15
End Sub
Private Sub Text1_GotFocus()
Call 全选
If Text1.Text = "<请输入ID>" Then Text1.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim sTemplate As String
    sTemplate = "!@#$%^&*()_+-=;,'.><\ /.][{}"    '用来存放不接受的字符
    If InStr(1, sTemplate, Chr(KeyAscii)) > 0 Then KeyAscii = 0: Call ShowMyTip(Text1, "错误", "请不要输入字符") Else TIP.Destroy
If KeyAscii = 22 Then KeyAscii = 0
If KeyAscii = 13 And Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then LogIn
End Sub

Private Sub Text1_LostFocus()
 TIP.Destroy
 If Len(Trim(Text1.Text)) = 0 Then Text1.Text = "<请输入ID>"
End Sub
Private Sub Text2_Change()
On Error Resume Next
Call ChangeValue(Text2)
End Sub

Private Sub Text2_DblClick()
Text2.SelLength = 0
Text2.SelLength = Len(Text2)
End Sub

Private Sub Text2_GotFocus()
Text2.SelLength = 0
Text2.SelLength = Len(Text2)
If Text2.Text = "<请输入密码>" Then Text2.Text = ""
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    SelPos = Text2.SelStart
    PwdLen = Len(Pwd)
    Insert = 0
    If KeyCode = 46 Then
If SelPos < PwdLen Then
    Pwd = Left(Pwd, SelPos) & Mid(Pwd, SelPos + 2)
    Call ChangeValue(Text2)
    KeyCode = 0
End If
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 22 Then KeyAscii = 0
If KeyAscii = 13 And Text1.Text <> "" And Text3.Text <> "" Then LogIn
    SelPos = Text2.SelStart
    PwdLen = Len(Pwd)
    Insert = 0

    Select Case KeyAscii
    Case 8
If SelPos > 0 Then
    Pwd = Left(Pwd, SelPos - 1) & Mid(Pwd, SelPos + 1)
    Insert = -1
End If
    Case 32 To 126
If (Text2.MaxLength > 0 And PwdLen < Text2.MaxLength) Or (Text2.MaxLength = 0) Then
    Pwd = Left(Pwd, SelPos) & Chr(KeyAscii) & Mid(Pwd, SelPos + 1)
    Insert = 1
End If
    End Select

    Call ChangeValue(Text2)
    KeyAscii = 0

End Sub

Private Sub Text2_LostFocus()
If Text2.Text = "" Then Text2.Text = "<请输入密码>"
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.SelLength = 0
Text2.SelStart = Len(Text2.Text)
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.SelLength = 0
Text2.SelStart = Len(Text2.Text)
End Sub

Private Sub Text2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.SelLength = 0
Text2.SelLength = Len(Text2)
End Sub

Private Sub Text3_GotFocus()
Call 全选
If Text3.Text = "<输入IP>" Then Text3.Text = ""
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Dim sTemplate As String
sTemplate = "!@#$%^&*()_+-=;,' ><\/.][{}ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"    '用来存放不接受的字符
If InStr(1, sTemplate, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Text3_LostFocus()
If Len(Trim(Text3.Text)) = 0 Then Text3.Text = "<输入IP>"
End Sub
Private Sub timetool_Timer()
TMRZ.Enabled = False
Timefriend.Enabled = False
If pl.Left > 5 Then
Call 反方向
Else
Call 正方向
End If
'反方向的意思是 从右往左 ←
End Sub
Sub 相册移动完成()

Timetool.Enabled = False
PF(6).Visible = True
PF(3).ScaleMode = 1
pl.Left = 0
If AUTOPLAYPIC = True Then
Timers.Enabled = True

End If
PicUse.Left = pl.Left - pl.Width
PICIM.Left = pl.Left + pl.Width
IMMAIN.Enabled = True
IMPIC.Enabled = True
IMCHAT.Enabled = True

End Sub
Sub 正方向()
If pl.Left >= -5 Then
Call 相册移动完成
Else
IMMAIN.Enabled = False
IMPIC.Enabled = False
IMCHAT.Enabled = False
PF(3).ScaleMode = 3
pl.Left = pl.Left + 15
PicUse.Left = pl.Left - PicUse.Width
PICIM.Left = pl.Left + pl.Width
If PICIM.Left >= PF(3).ScaleWidth Then Call 相册移动完成
End If
End Sub
Sub 反方向()
If pl.Left <= 5 Then
Call 相册移动完成
Else
IMMAIN.Enabled = False
IMPIC.Enabled = False
IMCHAT.Enabled = False
PF(6).Visible = True
PF(3).ScaleMode = 3
pl.Left = pl.Left - 15
PicUse.Left = pl.Left - PicUse.Width
PICIM.Left = pl.Left + pl.Width
If PicUse.Left <= -PF(3).ScaleWidth Then Call 相册移动完成
End If
End Sub
Private Sub timefriend_Timer()
Timetool.Enabled = False
TMRZ.Enabled = False
If PicUse.Left > -5 Then
IMMAIN.Enabled = True
IMPIC.Enabled = True
IMCHAT.Enabled = True
Timefriend.Enabled = False
PicUse.Left = 0
pl.Left = PicUse.Left + PicUse.Width
PICIM.Left = pl.Left + pl.Width
PF(6).Visible = False
Else
PF(3).ScaleMode = 3
IMMAIN.Enabled = False
IMPIC.Enabled = False
IMCHAT.Enabled = False
PF(6).Visible = True
PicUse.Left = PicUse.Left + 15
pl.Left = PicUse.Left + pl.Width
PICIM.Left = pl.Left + pl.Width
End If
End Sub
Private Sub Tit_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If H_CHANGE = True Then Exit Sub
Dim strpath As String
If Data.files.Count > 0 Then
strpath = Data.files(1)
Select Case LCase$(Right$(Data.files(1), 3))
Case "bmp", "dib", "gif", "jpg"
If fncGetInfo(strpath).PicHeight > 500 Or fncGetInfo(strpath).PicWidth > 500 Then
Call SHOWWRONG("图片像素过大，无法作为头像!", 2)
ElseIf fncGetInfo(strpath).PicHeight < 80 Or fncGetInfo(strpath).PicWidth < 80 Then
Call SHOWWRONG("图片像素过小，无法作为头像!", 2)
Else
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">成功更换了头像"
Call SaveSetting("ICEE", "Main", "logo", strpath)
Call LoadParam
End If
End Select
End If
End Sub

Private Sub TMAD_Timer()
On Error Resume Next
LOADTIME = LOADTIME + 1
If LOADTIME > 100 Then
TMAD.Enabled = False
Frmm.TMEA.Enabled = True
PDB.Left = PF(3).Left + PDB.Width
LOADTIME = 0
If FIRSTRUN = False Then Exit Sub
If ATP = 1 Then
sURL = GetInitEntry("PlayList", "LastUrl", "")  '初始化上次退出时播放的音乐路径
If Dir(sURL) = "" Then Exit Sub
Wm.URL = sURL
Wm.Controls.Play
If NEWS = 1 And Status.RasConnState = &H2000 Then Call DataNew '是否定制了ICEE每日资讯
End If
If ICK(2).Value = 1 And Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then Call LogIn '是否自动登录
End If
End Sub

Private Sub Tmrbk_Timer() '热键及其他
LA(3).Caption = PLIST.ListCount
LA(19).Caption = FRMDOWN.LVIEW.ListItems.Count
IW(7).SETTIP LA(19).Caption
LA(20).Caption = PLIST.ListCount
ML.Width = (PV.ScaleWidth * Wm.settings.volume) / 100
TMTIM.Enabled = PF(12).Visible

If PICAD.Visible = False Then Frmm.TMRCPU.Interval = 500 Else Frmm.TMRCPU.Interval = 100

If FAV_IT = True Then
If IMGFAV.PICTURE <> Frmm.PIC(52) Then IMGFAV.PICTURE = Frmm.PIC(52).PICTURE
Else
If IMGFAV.PICTURE <> Frmm.PIC(54) Then IMGFAV.PICTURE = Frmm.PIC(54).PICTURE
End If

If HAS_NET = False Then IMSIN.PICTURE = Frmm.PIC(32).PICTURE

Select Case WeekDay(Date, vbMonday)
Case 1
LA(17).Caption = "周一"
Case 2
LA(17).Caption = "周二"
Case 3
LA(17).Caption = "周三"
Case 4
LA(17).Caption = "周四"
Case 5
LA(17).Caption = "周五"
Case 6
LA(17).Caption = "周六"
Case 7
LA(17).Caption = "周日"
End Select

IW(8).SETTXT TimE
LA(25).Caption = Format(Now, "YYYY/MM/DD")
If Me.Enabled = True And QuickKey = 1 And PICAD.Visible = False Then '快捷键
If MyHotKey(vbKeyHome) Then Call iCan 'home显示
If MyHotKey(vbKeyEnd) And PICCLIP.Visible = False Then Call NoNoNo 'end隐藏
If MyHotKey(vbKeyPageUp) And Wm.playState = wmppsPlaying Then Call NT(1) 'PAGEUP上一首
If MyHotKey(vbKeyPageDown) And Wm.playState = wmppsPlaying Then 'PAGEDOWN下一首
If LOLIPOP <> 0 Then Call NT(2) Else Call NT(3) '下一首分两种情况
End If
If MyHotKey(vbKeyF8) Then Call 截屏 'F8截屏
If MyHotKey(vbKeyF10) And BACKME.Enabled = True Then frmset.Show   'F10 快速设置
End If

If IMCLEAR.Visible = True Then PMINFO.Left = IMCLEAR.Width + IMCLEAR.Left + 5 Else PMINFO.Left = IMCLEAR.Left

If TMP.Enabled = False Then  '右上角的时钟
Dim hourStr As String, minuteStr As String, secondStr As String
hourStr = Hour(TimE)
minuteStr = Minute(TimE)
secondStr = Second(TimE)
TimerStr(0) = IIf(Len(hourStr) = 2, Left(hourStr, 1), 0)
TimerStr(1) = IIf(Len(hourStr) = 2, Right(hourStr, 1), hourStr)
TimerStr(2) = IIf(Len(minuteStr) = 2, Left(minuteStr, 1), 0)
TimerStr(3) = IIf(Len(minuteStr) = 2, Right(minuteStr, 1), minuteStr)
TimerStr(4) = IIf(Len(secondStr) = 2, Left(secondStr, 1), 0)
TimerStr(5) = IIf(Len(secondStr) = 2, Right(secondStr, 1), secondStr)
PICTIME.Cls
LES = BitBlt(PICTIME.hdc, 0, 0, PICTIME.Width, PICTIME.Height, iFrame.hdc, PICTIME.Left, PICTIME.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\T" & TimerStr(0) & ".PNG", PICTIME.hdc, 16, 50)
Call PaintPng(App.Path & "\SKIN\T" & TimerStr(1) & ".PNG", PICTIME.hdc, 28, 50)
Call PaintPng(App.Path & "\SKIN\T" & TimerStr(2) & ".PNG", PICTIME.hdc, 48, 50)
Call PaintPng(App.Path & "\SKIN\T" & TimerStr(3) & ".PNG", PICTIME.hdc, 60, 50)
Call PaintPng(App.Path & "\SKIN\T" & TimerStr(4) & ".PNG", PICTIME.hdc, 79, 50)
Call PaintPng(App.Path & "\SKIN\T" & TimerStr(5) & ".PNG", PICTIME.hdc, 91, 50)
Call PaintPng(App.Path & "\SKIN\T_P.PNG", PICTIME.hdc, 40, 50)
Call PaintPng(App.Path & "\SKIN\T_P.PNG", PICTIME.hdc, 72, 50)
End If
PICTIME.Refresh
Dim r As RECT, p As POINTAPI, L As Long, rtn As Long, H As Long, H1 As Long, r1 As Long '鼠标移出/移入透明值得改变
L = GetWindowRect(Me.hwnd, r)
L = GetCursorPos(p)
GetCursorPos pos
SX = IIf(pos.X < 0 Or pos.X > Screen.Width / 15, IIf(pos.X < 0, 0, Screen.Width / 15), pos.X)
SY = IIf(pos.Y < 0 Or pos.Y > Screen.Height / 15, IIf(pos.Y < 0, 0, Screen.Height / 15), pos.Y)
If p.X < r.Left Or p.X > r.Right Or p.Y < r.Top Or p.Y > r.Bottom Then '移出界面
If TMO.Enabled = False Then Call MOVENOW
If frmmp.Visible = True And IS_M_S = True Then frmmp.Hide
If PF(15).Visible = True Then PF(15).Visible = False
If Pmusic.Visible = True Then Pmusic.Visible = False
End If
End Sub
Sub CPU() '绘制CPU波形图
On Error Resume Next
Dim lData As Long, r As Long
Dim hKey As Long, s As Long, LINEC As Long
s = 300
If once = True Then
Init
once = False
End If
Call PdhCollectQueryData(HQ)
r = CLng(PdhVbGetDoubleCounterValue(Counter, lData))
LA(9).Caption = Format$(r / 100, "##0#%")
If r > 0 And r <= 20 Then IMCPU.PICTURE = Frmm.PIC(0).PICTURE
If r > 20 And r <= 40 Then IMCPU.PICTURE = Frmm.PIC(7).PICTURE
If r > 40 And r <= 60 Then IMCPU.PICTURE = Frmm.PIC(4).PICTURE
If r > 60 And r <= 80 Then IMCPU.PICTURE = Frmm.PIC(25).PICTURE
If r > 80 And r <= 100 Then IMCPU.PICTURE = Frmm.PIC(9).PICTURE
LA(41).Caption = GetCPUTemp() & " °C"
Call RegOpenKey(HKEY_DYN_DATA, "PerfStats\StartStat", hKey)
Call RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, lType, lData, lSize)
Call RegCloseKey(hKey)
Call RegOpenKey(HKEY_DYN_DATA, "PerfStats\StatData", sk)
Call RegQueryValueEx(sk, "KERNEL\CPUUsage", 0, lType, lData, lSize)
a = a + 15
nx = a
ny = PCPU.Height - (r / 15) * PCPU.Height
PCPU.Line (Px, py + PCPU.Height * 12)-(nx, ny + PCPU.Height * 12)
Px = nx
py = ny
If a >= PCPU.ScaleWidth Then
a = 0
PCPU.Cls
nx = 0
ny = 0
Px = 0
py = 0
For LINEC = 1 To PCPU.ScaleHeight
PCPU.Line (LINEC * s, 0)-(LINEC * s, PCPU.ScaleHeight), vbBlack
Next
For LINEC = 1 To PCPU.ScaleWidth
PCPU.Line (0, LINEC * s)-(PCPU.ScaleWidth, LINEC * s), vbBlack
Next
End If
PCPU.Line (0, 0)-(PCPU.ScaleWidth - 15, PCPU.ScaleHeight - 15), COLOR_NOR, B
End Sub
Private Sub TMP_Timer()
'On Error Resume Next
Select Case Wm.playState '侦察状态
Case 6 '缓冲
LBSONG.Caption = "媒体缓冲中"
If D_L_SHOW = True Then FrmNetMusic.cDeskLrc.ShowText " ICEE音乐,音乐您的生活"
PICMU.Cls
PICMU.PICTURE = Frmm.da1.PICTURE
IW(0).SETEDIT False
Case 9 '连接
If songinx = 0 Then
LBSONG.Caption = "媒体连接中"
If D_L_SHOW = True Then FrmNetMusic.cDeskLrc.ShowText " ICEE音乐,音乐您的生活"
IW(0).SETEDIT False
PICMU.Cls
PICMU.PICTURE = Frmm.da1.PICTURE
Else
LBSONG.Caption = "超时重试中" & songinx
End If
Case 10 '错误
If songinx < 3 Then
LBSONG.Caption = "媒体打开错误"
songinx = songinx + 1
Wm.Controls.Stop
If D_L_SHOW = True Then FrmNetMusic.cDeskLrc.ShowText " ICEE音乐,音乐您的生活"
LBSONG.Caption = "还没有播放歌曲"
E2(2).Visible = False
E2(0).Visible = False
EI.Visible = False
TMP.Enabled = False
SB(2).Visible = False
PICMU.Cls
PICMU.PICTURE = Frmm.da1.PICTURE
IW(0).SETEDIT False
End If
Case 3 '播放
songinx = 0
If Len(SONGNAME) >= 35 Then LBSONG.Caption = Left(SONGNAME, 35) & "..." Else LBSONG.Caption = SONGNAME
SB(2).Visible = True
If Wm.currentMedia.durationString = "00:00/00:00" Then Call ORDERSONG '如果播放文件不正常则跳转下一首
If Wm.currentMedia.duration > 0 Then EI.Width = K(5).Width / Wm.currentMedia.duration * Wm.Controls.currentPosition  '如果进度条未被拖动 则计算出播放进度位置 并移动
If Running = False Then K(5).Visible = True
E2(2).Visible = True
E2(0).Visible = True
EI.Visible = True
Dim I As Integer, hourStr As String, minuteStr As String, secondStr As String
If UNCOUNT = False Then
hourStr = Hour(Wm.Controls.currentPositionString)
minuteStr = Minute(Wm.Controls.currentPositionString)
secondStr = Second(Wm.Controls.currentPositionString)
TimerStr(0) = IIf(Len(hourStr) = 2, Left(hourStr, 1), 0)
TimerStr(1) = IIf(Len(hourStr) = 2, Right(hourStr, 1), hourStr)
TimerStr(2) = IIf(Len(minuteStr) = 2, Left(minuteStr, 1), 0)
TimerStr(3) = IIf(Len(minuteStr) = 2, Right(minuteStr, 1), minuteStr)
TimerStr(4) = IIf(Len(secondStr) = 2, Left(secondStr, 1), 0)
TimerStr(5) = IIf(Len(secondStr) = 2, Right(secondStr, 1), secondStr)
Else
hourStr = Hour(Wm.currentMedia.durationString) - Hour(Wm.Controls.currentPositionString)
minuteStr = Minute(Wm.currentMedia.durationString) - Minute(Wm.Controls.currentPositionString)
secondStr = Second(Wm.currentMedia.durationString) - Second(Wm.Controls.currentPositionString)
TimerStr(0) = IIf(Len(hourStr) = 2, Left(hourStr, 1), 0)
TimerStr(1) = IIf(Len(hourStr) = 2, Right(hourStr, 1), hourStr)
TimerStr(2) = IIf(Len(minuteStr) = 2, Left(minuteStr, 1), 0)
TimerStr(3) = IIf(Len(minuteStr) = 2, Right(minuteStr, 1), minuteStr)
TimerStr(4) = IIf(Len(secondStr) = 2, Left(secondStr, 1), 0)
TimerStr(5) = IIf(Len(secondStr) = 2, Right(secondStr, 1), secondStr)
End If

PICTIME.Cls
LES = BitBlt(PICTIME.hdc, 0, 0, PICTIME.Width, PICTIME.Height, iFrame.hdc, PICTIME.Left, PICTIME.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\T" & TimerStr(0) & ".PNG", PICTIME.hdc, 16, 50)
Call PaintPng(App.Path & "\SKIN\T" & TimerStr(1) & ".PNG", PICTIME.hdc, 28, 50)
Call PaintPng(App.Path & "\SKIN\T" & TimerStr(2) & ".PNG", PICTIME.hdc, 48, 50)
Call PaintPng(App.Path & "\SKIN\T" & TimerStr(3) & ".PNG", PICTIME.hdc, 60, 50)
Call PaintPng(App.Path & "\SKIN\T" & TimerStr(4) & ".PNG", PICTIME.hdc, 79, 50)
Call PaintPng(App.Path & "\SKIN\T" & TimerStr(5) & ".PNG", PICTIME.hdc, 91, 50)
If UNCOUNT = True And TMP.Enabled = True Then Call PaintPng(App.Path & "\SKIN\UN_CO.PNG", PICTIME.hdc, 8, 50)
Call PaintPng(App.Path & "\SKIN\T_P.PNG", PICTIME.hdc, 40, 50)
Call PaintPng(App.Path & "\SKIN\T_P.PNG", PICTIME.hdc, 72, 50)
PICTIME.Refresh
PLAYDSB = PLAYDSB + 1
If PLAYDSB = 201 Then PLAYDSB = 197

If MAINSTYLE <> 3 And USE_PIC_FORM = False Then
PICMU.Cls
PICMU.PaintPicture Frmm.PIC(PLAYDSB).image, 0, 0, PICMU.ScaleWidth, PICMU.ScaleHeight
Else
Call IW(0).SETEDIT(True)
End If

E2(0).Width = EI.Width
SB(2).Width = EI.Width
Case 2 '暂停
PICMU.Cls
IW(0).SETEDIT False
PICMU.PICTURE = Frmm.da1.PICTURE
Case 1 '停止
PICMU.Cls
PICMU.PICTURE = Frmm.da1.PICTURE
EI.Visible = False
IW(0).SETEDIT False
LBSONG.Caption = "还没有播放歌曲"
If D_L_SHOW = True Then FrmNetMusic.cDeskLrc.ShowText " ICEE音乐,音乐您的生活"
K(5).Visible = False
EI.Visible = False
E2(0).Visible = False
E2(2).Visible = False
PLIST.ListIndex = Song
If Song < PLIST.ListCount - 1 Or LOLIPOP = 3 Or LOLIPOP = 1 Then
Call ORDERSONG
ElseIf LOLIPOP = 0 Then Call NT(3)
End If
End Select
End Sub

Sub ORDERSONG()
If LOLIPOP = 3 Then
PZOR.ToolTipText = "顺序播放"
LES = BitBlt(PZOR.hdc, 0, 0, PZOR.Width, PZOR.Height, PP.hdc, PZOR.Left, PZOR.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\SX_N.PNG", PZOR.hdc, 0, 0)
PZOR.Refresh
If PLIST.ListIndex < PLIST.ListCount - 1 Then  '如果列表中的歌还没有播放完
If rEADY = 1 Then
Song = Song + 1
Wm.URL = PLIST.URL(Song)
Wm.Controls.Play
rEADY = 2
ElseIf rEADY = 2 Then
Wm.Controls.Play
rEADY = 1
End If
ElseIf PLIST.ListIndex = PLIST.ListCount - 1 = True Then
Wm.Controls.Stop
PLIST.ListIndex = -1
End If
ElseIf LOLIPOP = 1 Then
PZOR.ToolTipText = "单曲循环"
LES = BitBlt(PZOR.hdc, 0, 0, PZOR.Width, PZOR.Height, PP.hdc, PZOR.Left, PZOR.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\DQ_N.PNG", PZOR.hdc, 0, 0)
PZOR.Refresh
If rEADY = 1 Then
  Wm.URL = LAST_URL
  Wm.Controls.Play
  rEADY = 2
ElseIf rEADY = 2 Then
  Wm.URL = LAST_URL
  Wm.Controls.Play
 rEADY = 1
End If

ElseIf LOLIPOP = 2 Then
PZOR.ToolTipText = "列表循环"
LES = BitBlt(PZOR.hdc, 0, 0, PZOR.Width, PZOR.Height, PP.hdc, PZOR.Left, PZOR.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\XH_N.PNG", PZOR.hdc, 0, 0)
PZOR.Refresh
If PLIST.ListIndex < PLIST.ListCount - 1 Then
If rEADY = 1 Then
Song = Song + 1
Wm.URL = PLIST.URL(Song)
Wm.Controls.Play
rEADY = 2
ElseIf rEADY = 2 Then
Wm.Controls.Play
rEADY = 1
End If
ElseIf PLIST.ListIndex = PLIST.ListCount - 1 Then
If rEADY = 1 Then
PLIST.ListIndex = 0
Wm.URL = PLIST.URL(0)
Wm.Controls.Play
rEADY = 2
ElseIf rEADY = 2 Then
Wm.Controls.Play
rEADY = 1
End If
End If
ElseIf LOLIPOP = 0 Then
PZOR.ToolTipText = "随机播放"
LES = BitBlt(PZOR.hdc, 0, 0, PZOR.Width, PZOR.Height, PP.hdc, PZOR.Left, PZOR.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\SJ_N.PNG", PZOR.hdc, 0, 0)
PZOR.Refresh
Call NT(3)
End If

lRet = SetInitEntry("Player", "zorder", LOLIPOP) '记录播放的循环模式
End Sub
Sub NT(Index As Integer) '上 / 下 曲控制 过程
On Error Resume Next
Dim URL As String, mm As String
Select Case Index
Case 1
If Song > 0 And PLIST.ListCount > 0 Then '上一曲
Song = Song - 1
PLIST.ListIndex = Song
URL = PLIST.URL(PLIST.ListIndex) '加载歌曲
If PLIST.ListIndex = 0 Then URL = PLIST.URL(PLIST.ListCount)
End If
Case 2
If LOLIPOP = 3 Or LOLIPOP = 1 Then
If Song < PLIST.ListCount - 1 Then '下一曲
Song = Song + 1
PLIST.ListIndex = Song
URL = PLIST.URL(PLIST.ListIndex) '加载歌曲
End If
End If
Case 3
Randomize
PLIST.ListIndex = Int(Rnd * (PLIST.ListCount - 1)) '随机时上/下 曲随机
URL = PLIST.URL(PLIST.ListIndex) '加载歌曲
End Select
If Trim(URL) = "" Then Exit Sub
Wm.URL = URL
Wm.Controls.Play '播放
mm = 0 '秒针归位
End Sub
Private Sub Timers_Timer()
If pl.Left = 0 Then Call SHOWU
End Sub
Private Sub TMO_Timer() '检测鼠标是否离开了窗体
If PICAD.Visible = True Then Exit Sub
Dim r As RECT, p As POINTAPI, L As Long
Dim rtn As Long
L = GetWindowRect(Me.hwnd, r)
L = GetCursorPos(p)
If p.X < r.Left Or p.X > r.Right Or p.Y < r.Top Or p.Y > r.Bottom And Me.Enabled = True Then
MakeTransparent Me.hwnd, 200
Call MOVENOW
Else
MakeTransparent Me.hwnd, 254
End If
End Sub


Private Sub TMRZ_Timer() '登陆框滑动
Timetool.Enabled = False
Timefriend.Enabled = False
If PICIM.Left < 5 Then
IMMAIN.Enabled = True
IMPIC.Enabled = True
IMCHAT.Enabled = True
Timers.Enabled = False
TMRZ.Enabled = False

PF(6).Visible = False
PICIM.Left = 0
pl.Left = PICIM.Left - pl.Width
PicUse.Left = pl.Left - PicUse.Width
Else
PF(3).ScaleMode = 3
IMMAIN.Enabled = False
IMPIC.Enabled = False
IMCHAT.Enabled = False
PF(6).Visible = True
PICIM.Left = PICIM.Left - 18
pl.Left = PICIM.Left - pl.Width
PicUse.Left = pl.Left - PicUse.Width
End If
End Sub

Private Sub tmrZoom_Timer()
Dim lRet    As Long
Dim ptMouse As POINTAPI
Static lElapsed As Long
        lElapsed = lElapsed + tmrZoom.Interval
        lRet = GetCursorPos(ptMouse)
        With ptMouse
            If (.X <> mlOldX) Or (.Y <> mlOldY) Or (lElapsed >= 250) Then
                Call DoZoom(ptMouse)
                lElapsed = 0
            End If
            mlOldX = .X
            mlOldY = .Y
        End With
End Sub

Private Sub TMTIM_Timer()
Dim hourStr As String, minuteStr As String
hourStr = Hour(Now)
minuteStr = Minute(Now)
TimerStr(0) = IIf(Len(hourStr) = 2, Left(hourStr, 1), 0)
TimerStr(1) = IIf(Len(hourStr) = 2, Right(hourStr, 1), hourStr)
TimerStr(2) = IIf(Len(minuteStr) = 2, Left(minuteStr, 1), 0)
TimerStr(3) = IIf(Len(minuteStr) = 2, Right(minuteStr, 1), minuteStr)
Dim I As Integer
For I = 0 To PTIME.Count - 1
PTIME(I).Cls
Next
Call PaintPng(App.Path & "\SKIN\BT" & TimerStr(0) & ".PNG", PTIME(0).hdc, 0, 0)
Call PaintPng(App.Path & "\SKIN\BT" & TimerStr(1) & ".PNG", PTIME(1).hdc, 0, 0)
Call PaintPng(App.Path & "\SKIN\BT" & TimerStr(2) & ".PNG", PTIME(3).hdc, 0, 0)
Call PaintPng(App.Path & "\SKIN\BT" & TimerStr(3) & ".PNG", PTIME(4).hdc, 0, 0)
Call PaintPng(App.Path & "\SKIN\BP.PNG", PTIME(2).hdc, 0, 0)

End Sub

Private Sub TreeView1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyDelete Then 移除好友
If KeyCode = vbKeyReturn Then 即时聊天
End Sub

Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then PopupMenu Frmm.mnuBuddy
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
'Debug.Print TreeView1.SelectedItem.Key
End Sub
Sub 列表搜索歌曲()
On Error Resume Next
PLIST.ListIndex = GetStringIndexInListBoxOrComboBox(Frmm.LISTM, TTS.Text, False)
PLIST.ListIndex = Frmm.LISTM.ListIndex
PLIST.Refresh
End Sub
Private Sub TTS_GotFocus()
If TTS.Text = "快速定位列表内歌曲" Then TTS.Text = ""
TTS.FOREColor = vbWhite
End Sub

Private Sub TTS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call 列表搜索歌曲
End Sub
Private Sub TTS_LostFocus()
TTS.Text = "快速定位列表内歌曲"
TTS.FOREColor = &HC0C0C0
End Sub
Private Sub TTS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
TTS.SetFocus
If Button = 2 Then Me.PopupMenu Frmm.文本
End Sub

'GET获取函数:参数:URL，返回值:源代码
Function XMLHttpGET(URL As String) As String
     Dim H As Object
     Set H = CreateObject("Microsoft.XMLHTTP")
     H.Open "GET", URL, True
     H.Send ""
     Do While H.ReadyState <> 4 '循环防止卡死
     DoEvents
     Loop
     If H.ReadyState = 4 Then XMLHttpGET = StrConv(H.responseBody, vbUnicode) '当h.ReadyState=4时说明源码加载完毕
 End Function

Private Sub TXTBAIDU_Change()
On Error Resume Next
Dim all As String, TMPE As String, Key As String
If Status.RasConnState <> &H2000 Then Exit Sub
Key = TXTBAIDU.Text
LISTBAIDU.Clear '每次文本框变化后都清除列表，以便加载新的数据
all = XMLHttpGET("http://suggestion.baidu.com/su.wd=" & GBtoUTF8(Key) & "&cb=window.bdsug.sug&from=superpage&t=1346931752233")
TMPE = Replace(Mid(all, InStr(all, "[") + 1, InStr(all, "]") - 36), """", "")   '截取数据
Dim a() As String, I As Integer
a = Split(TMPE, ",") '以分号隔开，分成数组
'循环将数组加载到列表
For I = 0 To UBound(a)
LISTBAIDU.AddItem a(I)
Next
If LISTBAIDU.ListCount > 3 Then
PICSER.Height = 190
LISTBAIDU.Visible = True
Else
PICSER.Height = 30
LISTBAIDU.Visible = False
End If
End Sub

Private Sub TXTBAIDU_GotFocus()
If TXTBAIDU.Text = "<请输入关键词>" Then TXTBAIDU.Text = ""
TXTBAIDU.FOREColor = vbWhite
If LISTBAIDU.ListCount > 3 Then
PICSER.Height = 190
LISTBAIDU.Visible = True
Else
PICSER.Height = 30
LISTBAIDU.Visible = False
End If
End Sub

Private Sub TXTBAIDU_KeyPress(KeyAscii As Integer)
Dim SERCH_ORDER As String
SERCH_ORDER = Trim(UCase(TXTBAIDU.Text))
If KeyAscii = 13 Then
Select Case SERCH_ORDER
Case "KILLME", "LETMEEND", "END"
Unload Me
Case "WHATISNEW"
FrmWhatNew.Show
Case "CHANGELOGO"
FRMHEAD.Show
Case "SCREENZOOMER"
Call SHOWZOOM
Case "Setting"
frmset.Show
Case "NEW", "NOTE"
Call LoadNote
Case "LOCKME"
Call Frmm.LOCKME
Case "UNSAFE"
Call ULock
Case "FAV_IT"
If frmma.Left > FRMFAV.Width Then
FRMFAV.Move frmma.Left - FRMFAV.Width, frmma.Top
Else
FRMFAV.Move frmma.Left + frmma.Width, frmma.Top
End If
FRMFAV.Show
Case Else
If Status.RasConnState <> &H2000 Then Exit Sub
ShellExecute 0&, vbNullString, URLTMP & TXTBAIDU.Text, vbNullString, vbNullString, 0 '调用ie
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">启动百度搜索引擎:" & TXTBAIDU.Text
End Select
End If
End Sub

Private Sub TXTBAIDU_LostFocus()
TXTBAIDU.FOREColor = &HE0E0E0
If LISTBAIDU.ListCount = 0 Then PICSER.Height = 30
If Trim(TXTBAIDU.Text) = "" Then TXTBAIDU.Text = "<请输入关键词>"
End Sub

Private Sub TXTBAIDU_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMBAIDU.PICTURE <> Frmm.X1.PICTURE Then IMBAIDU.PICTURE = Frmm.X1.PICTURE

End Sub

Private Sub TXTBOX_Change()
    Dim I As Long, j As Long
    Dim strPartial As String, strTotal As String
    If m_bEditFromCode Then
m_bEditFromCode = False
Exit Sub
    End If
    With TXTBOX
strPartial = .Text
I = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)
If I <> CB_ERR Then
    strTotal = .List(I)
    j = Len(strTotal) - Len(strPartial)
    '
If j <> 0 Then
      
m_bEditFromCode = True
.SelText = Right$(strTotal, j)
.SelStart = Len(strPartial)
.SelLength = j
    Else

    End If
End If
    End With
End Sub
Private Sub TXTBOX_KeyPress(KeyAscii As Integer)
Dim sTemplate As String
    sTemplate = "!@#$%^&*()_+-=,.'<>;/\[.]{}"   '用来存放不接受的字符
    If InStr(1, sTemplate, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtDecimal_KeyPress(KeyAscii As Integer)
KeyAscii = VailText(KeyAscii, "0123456789", True)
End Sub
Private Sub txtEntry_Change()
    If txtEntry.Text = "+" Or txtEntry.Text = "*" Or txtEntry.Text = "/" Then
        txtEntry.Text = "ans" + txtEntry.Text
        txtEntry.SelStart = Len(txtEntry.Text)
    End If
End Sub

Private Sub txtEntry_GotFocus()
If PF(9).Visible = True Then PF(9).Visible = False
End Sub

Private Sub txtEntry_KeyPress(KeyAscii As Integer)
If txtEntry.Text = "" Then Exit Sub
If KeyAscii = 13 Then Call CalculateEntry: txtEntry.Text = ""
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
Dim sTemplate As String
    sTemplate = "!@#$%^&*()_+-=;,. '><\/.][{}ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"    '用来存放不接受的字符
    If InStr(1, sTemplate, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtFrom_LostFocus()
If Trim(txtFrom.Text) = "" Then txtFrom.Text = "0"
End Sub

Private Sub txtLogBase_KeyPress(KeyAscii As Integer)
KeyAscii = VailText(KeyAscii, "0123456789.-", True)
End Sub



Private Sub TXTPOUP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then READPASS
End Sub




Private Sub TXTSER_GotFocus()
If TXTSER.Text = "请输入对方ID" Then TXTSER.Text = ""
End Sub

Private Sub TXTSER_KeyPress(KeyAscii As Integer)
Dim sTemplate As String
sTemplate = "!@#$%^&*()_+-=,;'. \/][><.{}"   '用来存放不接受的字符
If InStr(1, sTemplate, Chr(KeyAscii)) > 0 Then KeyAscii = 0
If KeyAscii = 13 Then ADDFRIEND
End Sub

Private Sub TXTSER_LostFocus()
If Len(Trim(TXTSER.Text)) = 0 Then TXTSER.Text = "请输入对方ID"
End Sub

Private Sub TXTSER_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicUse_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub
Private Sub txtText_Change()
On Error Resume Next
Open App.Path & "\COFING\CLIPTEXT.txt" For Binary As #1
Put #1, LOF(1) + 1, Now & vbCrLf & txtText.Text & vbCrLf
Close #1
End Sub

Private Sub txtText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtText.SetFocus: If Button = 2 Then Me.PopupMenu Frmm.文本, 0
End Sub


Private Sub txtText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If txtText.BackColor <> vbWhite Then txtText.BackColor = vbWhite
End Sub

Private Sub txtTo_Change()
If Len(txtTo) <> 0 Then lRet = SetInitEntry("SYSTEM", "NETTO", txtTo.Text)
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
Dim sTemplate As String
sTemplate = "!@#$%^&*()_+-=;,.' ><\/.][{}ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"    '用来存放不接受的字符
If InStr(1, sTemplate, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtTo_LostFocus()
If Trim(txtTo.Text) = "" Then txtTo.Text = "99"
End Sub
Public Sub LoadSettings()
On Error Resume Next
Dim LINEC As Long
For LINEC = 1 To PCPU.ScaleHeight
PCPU.Line (LINEC * 300, 0)-(LINEC * 300, PCPU.ScaleHeight), vbBlack
Next
For I = 1 To PCPU.ScaleWidth
PCPU.Line (0, LINEC * 300)-(PCPU.ScaleWidth, LINEC * 300), vbBlack
Next
IS_AM = GetInitEntry("SYSTEM", "AM", True)
LOLIPOP = GetInitEntry("Player", "zorder", 0) '初始化播放顺序
UNCOUNT = GetInitEntry("PLAYER", "TIMEMODE", False)
AUTOSERCH = GetInitEntry("PLAYER", "AUTOSERCHLRC", 1)
GETWEATHER = GetInitEntry("System", "Weather", 0) '自动下载'原为天气预报，因代码原因更改
ALWAYSONTOP = GetInitEntry("SYSTEM", "ONTOP", False)
If LOLIPOP = 3 Then
PZOR.ToolTipText = "顺序播放"
Call PaintPng(App.Path & "\SKIN\SX_N.PNG", PZOR.hdc, 0, 0)
ElseIf LOLIPOP = 1 Then
PZOR.ToolTipText = "单曲循环"
Call PaintPng(App.Path & "\SKIN\DQ_N.PNG", PZOR.hdc, 0, 0)
ElseIf LOLIPOP = 2 Then
PZOR.ToolTipText = "列表循环"
Call PaintPng(App.Path & "\SKIN\XH_N.PNG", PZOR.hdc, 0, 0)
ElseIf LOLIPOP = 0 Then
PZOR.ToolTipText = "随机播放"
Call PaintPng(App.Path & "\SKIN\SJ_N.PNG", PZOR.hdc, 0, 0)
End If
USE_PIC_FORM = GetInitEntry("SYSTEM", "USE_PIC", True)
USEBACK = GetInitEntry("SYSTEM", "BACKPICTURE", App.Path + "\SKIN\BK\0.JPG")
P_BK_INDEX = GetInitEntry("SYSTEM", "BACKPICTURE_INDEX", 0)
IMSIGN.Move MBK(P_BK_INDEX).Left + MBK(P_BK_INDEX).Width - IMSIGN.Width, MBK(P_BK_INDEX).Top
Frmm.IMBK.PICTURE = LoadPicture(USEBACK)
HAS_HEAD = GetInitEntry("SYSTEM", "HEAD_VIS", True)
lRet = SetInitEntry("Time", "Start", Now) '记录启动时间
lRet = SetInitEntry("Path", "App.Path", App.Path) '记录启动路径
lRet = SetInitEntry("Pid", "Pid", App.ThreadID)   '记录程序ID
lRet = SetInitEntry("Version", "Version", App.Revision) '记录程序版本(为检查更新做准备)
lRet = SetInitEntry("User", "IPADDRESS", GetIPAddress)  '记录本地ip地址
lRet = SetInitEntry("User", "PCNAME", GetIPHostName)    '记录计算机名称
lRet = SetInitEntry("User", "LOGO", LOGO)   '记录本地头像的路径

SCREENSAVER = GetInitEntry("ScreenSaver", "Opened", 1)     '初始化是否在程序运行期间允许屏保的运行
QuickKey = GetInitEntry("System", "Quickkey", 0) '快捷键
MOVE_TRANS = GetInitEntry("SYSTEM", "transparent", 0)
If MOVE_TRANS = 1 Then TMO.Enabled = True Else TMO.Enabled = False: MakeTransparent Me.hwnd, 254
R_P_THU = GetInitEntry("SYSTEM", "REPLACE", 0)
txtTo.Text = GetInitEntry("SYSTEM", "NETTO", 99)
ATP = GetInitEntry("Player", "AutoPlay", 0)    '初始化是否自动播放列表
ICK(1).Value = GetInitEntry("IM", "RememberPassWord", 1)  '初始化是否记录密码
ICK(2).Value = GetInitEntry("IM", "AutoLogin", 0) '自动登录
Text1.Text = GetInitEntry("IM", "LastUserID", "<请输入ID>") '加载ID
Text3.Text = GetInitEntry("IM", "LastServerIp", Winsock1.LocalIP) '加载服务地址
Pwd = GetInitEntry("IM", "LastPassWord", "")
If ICK(1).Value = 1 Then Text2.Text = Pwd '加载密码
Set mmp = New clsNet
Let myIp = Winsock1.LocalIP
Let netIp = Get_Net_Ip(myIp)
Let lblFrom.Caption = "起始 " & netIp
Let lblTo.Caption = "结束 " & netIp
If ALWAYSONTOP = True Then RESL = SetWindowPos(frmmabk.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags) Else RESL = SetWindowPos(frmmabk.hwnd, 1, 0, 0, 0, 0, flags)
If ALWAYSONTOP = True Then RESL = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags) Else RESL = SetWindowPos(Me.hwnd, 1, 0, 0, 0, 0, flags)
For I = 0 To IST.Count - 1
IST(I).IS_SELECT = False
Next
IST(MAINSTYLE).IS_SELECT = True
GETMSGCOUNT = GetInitEntry("IM", "MSGCOUNT", 0)
LCO.Caption = GETMSGCOUNT
PF(4).BackColor = GetInitEntry("SYSTEM", "WIN_COLOR", vbBlack)
ICOCO(GetInitEntry("SYSTEM", "WIN_COLOR_SEL", 0)).IS_SELECT = True
End Sub
Private Sub SaveSettings() '保存配置
lRet = SetInitEntry("Time", "Last End", Now) '记录关闭时间
lRet = SetInitEntry("Settings", "Left", Me.Left) '记录窗体位置
lRet = SetInitEntry("Settings", "Top", Me.Top)  '同上
Call SaveSetting("ICEE", "MAIN", "top", Me.Top) '保存窗体位置
Call SaveSetting("ICEE", "MAIN", "Left", Me.Left) '保存窗体位置
lRet = SetInitEntry("Player", "Volume", Me.Wm.settings.volume)   '记录音乐播放器的音量
lRet = SetInitEntry("Player", "zorder", LOLIPOP) '记录播放的循环模式
lRet = SetInitEntry("IM", "UseNewUser", ICK(0).Value) '记录是否新用户
lRet = SetInitEntry("IM", "RememberPassWord", ICK(1).Value) '记住密码
lRet = SetInitEntry("IM", "LastUserID", Text1.Text)   '记录最后登录的用户ID
lRet = SetInitEntry("IM", "LastServerIp", Text3.Text) '记录最后的服务器地址
lRet = SetInitEntry("SYSTEM", "COLOR", PicUse.BackColor) '记录主界面的背景颜色
lRet = SetInitEntry("PLAYER", "TIMEMODE", UNCOUNT) '保存是否启动音乐播放的倒计时
End Sub
Sub LoadPic()
On Error Resume Next
Dim lIdx    As Long
Dim lPicCnt As Long
Dim lFilCnt As Long
picSlide.Visible = False
LA(24).Caption = filHidden.Path
picSlide.Move 0, 64, IPLAY(0).Width, IPLAY(0).Height
    While IPLAY.Count > 1
        Unload IPLAY(IPLAY.Count - 1)
    Wend
    lFilCnt = filHidden.ListCount
        For lIdx = 0 To filHidden.ListCount - 1
            ERR.Clear
                If lPicCnt > 0 Then
                    Load IPLAY(lPicCnt)
                    Set IPLAY(lPicCnt).Container = picSlide
                    End If
                IPLAY(lPicCnt).AUTOSIZE = False
                IPLAY(lPicCnt).HASTIP = False
                IPLAY(lPicCnt).SETTIP filHidden.Path & "\" & filHidden.List(lIdx)
                IPLAY(lPicCnt).IS_PIC = True
                IPLAY(lPicCnt).SETPIC filHidden.Path & "\" & filHidden.List(lIdx)
                IPLAY(lPicCnt).Visible = True
                lPicCnt = lPicCnt + 1
        Next lIdx
        Call PIC_RESIZE
        picSlide.Visible = True
        LA(27).Caption = filHidden.ListCount
End Sub
Sub PIC_RESIZE()
Dim X       As Long
Dim Y       As Long
Dim lIdx    As Long
            For lIdx = 0 To IPLAY.Count - 1
                X = lIdx * IPLAY(0).Width
                Y = 0
                IPLAY(lIdx).Move X, Y
            Next lIdx
            picSlide.Width = lIdx * IPLAY(0).Width
            HScroll1.Value = 0
            HScroll1.MaxV = picSlide.Width - Picture1.ScaleWidth
            
            If HScroll1.MaxV < 0 Then
                HScroll1.MaxV = 0
            Else
                HScroll1.SmallChange = IPLAY(0).Width
                HScroll1.LargeChange = Picture1.ScaleWidth
            End If
End Sub
Private Sub HScroll1_Change()
picSlide.Left = -HScroll1.Value
End Sub
Private Function CreateCheckeredBrush(ByVal hdc As Long, ByVal lColor1 As Long, ByVal lColor2 As Long) As Long
Dim X As Long
Dim Y As Long
Dim lRet As Long
Dim hBitmapDC As Long
Dim hBitmap As Long
Dim hOldBitmap As Long
If lColor1 < 0 Then
lColor1 = GetSysColor(lColor1 And &HFF&)
End If
If lColor2 < 0 Then
lColor2 = GetSysColor(lColor2 And &HFF&)
End If
hBitmapDC = CreateCompatibleDC(hdc)
hBitmap = CreateCompatibleBitmap(hdc, 8, 8)
hOldBitmap = SelectObject(hBitmapDC, hBitmap)
For Y = 0 To 6 Step 2
For X = 0 To 6 Step 2
lRet = SetPixelV(hBitmapDC, X, Y, lColor1)
lRet = SetPixelV(hBitmapDC, X + 1, Y, lColor2)
lRet = SetPixelV(hBitmapDC, X, Y + 1, lColor2)
lRet = SetPixelV(hBitmapDC, X + 1, Y + 1, lColor1)
Next
Next
hBitmap = SelectObject(hBitmapDC, hOldBitmap)
CreateCheckeredBrush = CreatePatternBrush(hBitmap)
lRet = DeleteDC(hBitmapDC)
lRet = DeleteObject(hBitmap)
End Function
Sub DisableCtrlAltDelete(bDisabled As Boolean) '禁止都打开任务管理器
Dim X As Long
X = SystemParametersInfo(97, bDisabled, CStr(1), 0)
End Sub
Sub DrawGray() '转换为灰度图像
Dim Red As Integer
Dim Green As Integer
Dim Blue As Integer
Dim C As Long
Dim graycolor As Long
Dim X0 As Integer
Dim Y0 As Integer
For X0 = 0 To Picture1.Width
For Y0 = 0 To Picture1.Height
C = Picture1.POINT(X0, Y0)
Red = (C And &HFF)
Green = (C And 62580) / 256
Blue = (C And &HFF00) / 65536
graycolor = (Red + Green + Blue) / 3
Picture2.PSet (X0, Y0), RGB(graycolor, graycolor, graycolor)
Next
Next
PLOGU.PaintPicture Picture2.image, 0, 0, PLOGU.Width, PLOGU.Height
Call PaintPng(App.Path & "\SKIN\HEAD.png", PLOGU.hdc, 0, 0 - 1)
End Sub
Sub LoadList() '加载列表
Call Playlist(App.Path + "\Media\plist.m3u")
PLIST.Refresh

End Sub
Sub checkSound() '检测是否安装了声卡
Dim I As Integer
    I = waveOutGetNumDevs()
    If I > 0 Then
    USELOGO.Enabled = True
    LBITEM(7).Enabled = True
    Else
    USELOGO.Enabled = False
    LBITEM(7).Enabled = False
    End If
    If Left(IEver, 1) < 7 Then Debug.Print "请安装IE7或更高版本"
End Sub
Private Sub Form_Activate()
On Error Resume Next
Dim myUsage As Double
If ALWAYSONTOP = True Then RESL = SetWindowPos(frmmabk.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags) Else RESL = SetWindowPos(frmmabk.hwnd, 1, 0, 0, 0, 0, flags)
If ALWAYSONTOP = True Then RESL = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags) Else RESL = SetWindowPos(Me.hwnd, 1, 0, 0, 0, 0, flags)
myUsage = RamUsage
LA(23).Caption = FormatUsage(myUsage) & "K (" & FormatUsage(myUsage / 1024) & " Mb)"
myUsage = PFUsage
LA(32).Caption = FormatUsage(myUsage) & "K (" & FormatUsage(myUsage / 1024) & " Mb)"
Call UnHook
H_DOS = 7
gHW = Me.hwnd '鼠标控件
Call Hook '唤醒鼠标滑轮API
PSubClass.AddWindowMsgs Me.hwnd
Call MoveWindow(frmmabk.hwnd, Me.Left / Screen.TwipsPerPixelX - 20, Me.Top / Screen.TwipsPerPixelY - 10, 380, 650, True)
Call Frmm.CHECKNET
If Status.RasConnState = &H2000 Then HAS_NET = True Else HAS_NET = False
Call RUNSAVER
End Sub
Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance Then End
If Sound = 1 Then sndPlaySound App.Path + "\Sound\Load.wav", 1        '播放登陆声音
Call LoadParam
Me.Hide
FILESINGER = ""
Load FRMFAV
PICAD.Visible = True
iFrame.Move 1, 1
PF(11).Move 0, 0, iFrame.ScaleWidth, iFrame.ScaleHeight
Frmm.WB.Navigate "http://hi.baidu.com/iceeorgan/item/96d45007a86c1acbff240dfa"
Dim Buffer As String, LST As String 'declare the needed variables
Buffer = Space(MAX_PATH)
rtn = GetSystemDirectory(Buffer, Len(Buffer)) 'get the path
rtn = GetWindowsDirectory(Buffer, Len(Buffer)) 'get the path
WinSysPath = Left(Buffer, rtn)                'parse the path into the global string
WinPath = Left(Buffer, rtn)                    'parse the path to the global string
lRet = SetInitEntry("OS", "WINPATH", WinPath)
lRet = SetInitEntry("OS", "WINSYSPATH", WinSysPath)

  Select Case Format(Now, "YYYY") Mod 12
    Case 4
      LA(26).Caption = "鼠年"
    Case 5
      LA(26).Caption = "牛年"
    Case 6
      LA(26).Caption = "虎年"
    Case 7
      LA(26).Caption = "兔年"
    Case 8
      LA(26).Caption = "龙年"
    Case 9
      LA(26).Caption = "蛇年"
    Case 10
     LA(26).Caption = "马年"
    Case 11
      LA(26).Caption = "羊年"
    Case 0
     LA(26).Caption = "猴年"
    Case 1
      LA(26).Caption = "鸡年"
    Case 2
     LA(26).Caption = "狗年"
    Case 3
     LA(26).Caption = "猪年"
   End Select

IW(11).SETTIP "图库"
IW(11).SETTXT ""
IW(11).SETPATH GetInitEntry("DESK", "PICSHOW", App.Path + "\MEDIA\PIC")
IW(11).IS_PIC_SHOW = True
Me.Top = GetSetting("ICEE", "MAIN", "Top", (Screen.Height - Me.Height) / 2)
Me.Left = GetSetting("ICEE", "MAIN", "Left", (Screen.Width - Me.Width) / 2)
If Me.Top > Screen.Height Then Me.Top = 0
If Me.Left > Screen.Width Then Me.Left = 0
Call SaveSetting("ICEE", "Winsock", "Connect", 0) '启动时应该先初始化注册表，0时为未登录
'为程序设置设置运行等级

Set PSubClass = New cSubclass '继承无拖影
Call PSubClass.AddWindowMsgs(Me.hwnd)  '继承无拖影

IS_CHECK_CLIP = GetInitEntry("SYSTEM", "CHECK_CLIP", True)

IW(8).HASTIP = False

Dim HProcress As Long, PPriorityClass As Long '
HProcress = OpenProcess(PROCESS_ALL_ACCESS, 0, GetCurrentProcessId)
SetPriorityClass HProcress, NORMAL_PRIORITY_CLASS '标准的运行等级
'屏蔽右键菜单
IW(8).IS_PIC = False
IW(8).MY_STYLE = 1
 IW(8).SETFONT "微软雅黑", 14, True, 12, False
 IW(7).SETFONT "微软雅黑", 12, True, 16, False
lRet = SetInitEntry("MsgBOX", "LEFT", frmma.Left)
lRet = SetInitEntry("MsgBOX", "TOP", frmma.Top + (frmma.Height - 4000) / 2)
IETIP = GetInitEntry("SYSTEM", "IETIP", 0)
lbthing.Caption = "欢迎使用1.24全新版本,更多精彩等你发现"
Load FRMLRC
Dim TXTBOX As Control
For Each TXTBOX In Me.Controls
If TypeOf TXTBOX Is TextBox Then
oldproc = GetWindowLong(TXTBOX.hwnd, GWL_WNDPROC)
SetWindowLong TXTBOX.hwnd, GWL_WNDPROC, AddressOf TextWndProc
End If
Next

SHRO.Value = 0
SHRO.Max = PF(4).Height - PicUse.ScaleHeight

filHidden.Path = GetInitEntry("PHOTO_PLAYER", "PATH", App.Path & "\SKIN\PHOTO")
filHidden.Pattern = "*.BMP;*.JPG"

RPC.ROUND_PIC PSEND, 4, 0, 0
RPC.ROUND_PIC PMDL, 4, 0, 0

PICMU.PICTURE = Frmm.da1.PICTURE
E2(2).Move 0, 153
E2(0).Move 0, 153
lbthing.Move 0, 8

PNZ.PICTURE = Frmm.da2.image
IMGUSB.PICTURE = Frmm.PIC(37).PICTURE
ML.Width = GetInitEntry("PLAYER", "VOLOME", PV.ScaleWidth)
Wm.settings.volume = Int((100 / PV.ScaleWidth) * ML.Width) '初始化播放器音量
'移动控件位置
PP.Move 0, 110, 340, 441
PF(0).Move 0, 185, 340, 367
PF(3).Move 0, 185, 340, 367 '主要功能区
PF(4).Move 0, 0
PF(6).Move 0, 0
Pmusic.Move 0, 0, 340, 355
PicZoom.AutoRedraw = True
Winsock1.RemotePort = "6000"
Winsock1.Listen

PICD.Move 0, 155, 340, 400 '变更图片区
TXTSER.Text = "请输入对方ID"
LBSONG.Caption = "还没有播放歌曲"
LBITEM(3).Caption = "主菜单"
Wm.Controls.Stop
LA(3).FontName = "微软雅黑"
Px = 0 '这句是CPU绘图获取宽度
py = PCPU.Height '这句是CPU绘图获取高度
pl.Move 0, 0, PF(3).ScaleWidth, PF(3).ScaleHeight
PicUse.Move 0, 0, PF(3).ScaleWidth, PF(3).ScaleHeight
PICIM.Move 0, 0, PF(3).ScaleWidth, PF(3).ScaleHeight
PICLO.Move 0, 0, PF(3).ScaleWidth, PF(3).ScaleHeight
PICLO.ZOrder 0
E2(0).Width = 1

pl.Left = PicUse.Left + pl.Width
PICCLIP.Move 0, 0, 340
PICNET.Move 0, 26, 340
PICNET.Move 0, 26, 340
PICBUG.Move 0, 30, 340
PICFI.Move 0, 30, 340
PICIG.Move 0, 30, 340
PICPASS.Move 0, 30, 340
Dim SKU As Integer
For SKU = 0 To LBITEM.Count - 1
LBITEM(SKU).OLEDropMode = 1
LBITEM(SKU).FOREColor = vbWhite
Next
For SKU = 0 To IW.Count - 1
IW(SKU).HASLINE = False
Next

For SKU = 0 To PMU.Count - 1
PMU(SKU).HASLINE = False
PMU(SKU).HASTIP = True
PMU(SKU).IS_PIC = True
Next
PMU(0).SETCOLOR Frmm.PIC(3).POINT(0, 0), Frmm.PIC(3).POINT(0, 0)
PMU(1).SETCOLOR Frmm.PIC(22).POINT(0, 0), Frmm.PIC(22).POINT(0, 0)
PMU(2).SETCOLOR Frmm.PIC(28).POINT(0, 0), Frmm.PIC(28).POINT(0, 0)
PMU(3).SETCOLOR Frmm.PIC(23).POINT(0, 0), Frmm.PIC(23).POINT(0, 0)
SURO.MaxV = PF(16).Height - PF(14).ScaleHeight + 15
SURO.LargeChange = 15
SURO.Value = 0

PMU(0).SETIMG Frmm.PIC(3)
PMU(1).SETIMG Frmm.PIC(22)
PMU(2).SETIMG Frmm.PIC(28)
PMU(3).SETIMG Frmm.PIC(23)
PMU(4).SETIMG Frmm.PIC(30)

PMU(0).SETTIP "收藏夹"
PMU(1).SETTIP "电台"
PMU(3).SETTIP "本地"
PMU(4).SETTIP "播放历史"
PMU(2).SETTIP "添加"

For SKU = 0 To ICK.Count - 1
ICK(SKU).SETCOLOR &H383537, vbWhite
ICK(SKU).M_STYLE = 2
Next
For I = 0 To IWG.Count - 1
IWG(I).IS_PIC = True
IWG(I).SETCOLOR COLOR_NOR, COLOR_HIGH
IWG(I).HASLINE = False
Next
ICK(0).SETTXT "新用户登录"
ICK(1).SETTXT "记住密码"
ICK(2).SETTXT "自动登录"
ICL(4).SETTXT "退出ICEE"
ICL(0).SETTXT "解除锁定"
ICL(2).M_STYLE = 1
ICL(2).SETTXT "是,立刻升级"
ICL(3).SETTXT "不了,下次别提醒我"
ICL(4).SETTXT "详情"
ICL(5).SETTXT "更换相册"
ICL(6).SETTXT "取消"
ICL(7).SETTXT "扫描"
ICL(1).SETTXT "登陆"

ICL(0).SETCOLOR &HBCD3E2, vbWhite, vbBlack
ICL(2).SETCOLOR &H231C09, &H899F1E, vbWhite
ICL(3).SETCOLOR &H231C09, &H899F1E, vbWhite
ICL(4).SETCOLOR &HE3E3E3, vbWhite, vbBlack
ICL(1).SETCOLOR &H373637, &H899F1E, vbWhite
ICL(7).SETCOLOR &H373637, &H899F1E, vbWhite
IST(0).SETTXT "iOS图标"
IST(1).SETTXT "卡通图标"
IST(2).SETTXT "圆角图标"
IST(3).SETTXT "Win8主题"
ICS(0).IS_SELECT = True
PF(15).Visible = True
ICS(0).SETTXT "背景色"
ICS(1).SETTXT "瓷砖颜色"
ICC(0).SETTXT "保存"
ICC(1).SETTXT "编辑"
ICC(2).SETTXT "打印"
ICC(3).SETTXT "分享"
For SKU = 0 To ICC.Count = 1
ICC(I).HASLINE = False
Next
ICT.SETTXT "查看剪切板历史文本"

PF(2).Line (0, 0)-(PF(2).ScaleWidth - 1, PF(2).ScaleHeight - 1), &H808080, B
PF(12).Move 0, 0, iFrame.ScaleWidth, iFrame.ScaleHeight
LA(4).Caption = GetInitEntry("SCREEN_MAKER", "ZOOMPER", 100)
ICZ(3).SETTXT "←"
mlOldX = -100
mfScale = LA(4).Caption / 100!
PicZoom.Move 0, 0, iFrame.ScaleWidth, iFrame.ScaleHeight

PF(15).Move 0, 105
PICTOOL.Move 0, 551
PR.Move 0, PICTOOL.Top - PR.Height
PF(2).Move 0, PICTOOL.Top - PF(2).Height
PICSER.Move 0, 155, iFrame.ScaleWidth, 30
'初始化播放器控件
PICBACK.Move 0, 0 '列表按钮位置
PSEND.Move 8, 104
PMDL.Move 64, 104
PICDL.Move 264, 110 '下载按钮位置
IU.Move 0, 0 '关闭按钮位置
IPRE.Move 56, 376 '上一首按钮
INEXT.Move 232, 376 '下一首按钮
PLAYB.Move 128, 352 '播放按钮
IMVOL.Move 300, 384 '音量
PV.Move 8, 384
Pmusic.Visible = False '列表不可见
PICBACK.Visible = True '返回按钮
IU.Visible = True
PICDL.Visible = True
LBSONG.Move 40, 56 '歌名标签
K(5).Move 0, 439, 340 '进度背景
EI.Move 0, 439  '进度条
Wm.Move 999, 999
PLIST.Move 0, 40, Pmusic.ScaleWidth, Pmusic.ScaleHeight - Mbar.Height - PLIST.Top
Frmm.PSINGER.PICTURE = Frmm.PIC(152).image
PLAYDSB = 197
Wm.Visible = False
Timers.Enabled = False
ListView1.ColumnHeaders.Add , , "文件类型", 280
Pser.Move 0, Mbar.Top - Pser.Height
IMJ.Visible = False '关闭子界面按钮
SetTrayIcon Frmm.OFFLINE.PICTURE
Call setDefaultConfig '默认信息
Call KillAuto(GetUdisk) '删除磁盘autorun文件
Call UpdateDisk '更新磁盘设备
Call LoadPic  '初始化图片
Call ShowFriend '主菜单呈现
Call 初始化
Call LoadSettings '加载配置文件
Call SeekMe(Me)  '不让窗体跑出桌面
Call SetSEH(True) '为了避免程序在中途崩溃的样子太丑，此函数将会提醒用户终止程序，放弃程序，忽略的操作
Call checkSound '检查声卡
Call SaveSetting("ICEE", "Path", "Path", App.Path) '保存程序安装路径
Select Case WeekDay(Date, vbMonday)
Case 1
THIS_DAY = Date & " - 星期" & "一"
Case 2
THIS_DAY = Date & " - 星期" & "二"
Case 3
THIS_DAY = Date & " - 星期" & "三"
Case 4
THIS_DAY = Date & " - 星期" & "四"
Case 5
THIS_DAY = Date & " - 星期" & "五"
Case 6
THIS_DAY = Date & " - 星期" & "六"
Case 7
THIS_DAY = Date & " - 星期" & "日"
End Select
iCount = 0
NOTECOUND = 0 '便签数量
once = True
Running = False
MINSOUND = False
FIRSTRUN = True
IS_FIRST_LOAD_ACT = True
rEADY = 1
MB = 0
Set CLIPS = New Collection
Set objTimer = New clsWaitableTimer
Call SubClass(Me.hwnd)

If IS_CHECK_CLIP = True Then
Call StartMonitoring(Me.hwnd)
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">启动了剪切板监视"
Else
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">关闭了剪切板监视"
End If
Call DRAWFACE
memory.dwLength = Len(memory)
Call GlobalMemoryStatus(memory) '获得内存信息
PDB.Move 0, 155, 340, 396
IMJ.ZOrder 0
PICAD.ZOrder 0
PICAD.Move 0, 0, iFrame.ScaleWidth, iFrame.ScaleHeight
gFileNum = FreeFile
PR.Visible = False
IMJ.Visible = False
Call GetDriver '获取U盘.
Call LoadList '列表
Call Frmm.CHECKNET
MakeTransparent Me.hwnd, 254
Call iCan
sIndex = GetInitEntry("Playlist", "LastIndex", 0) '获得上次退出时播放曲目的索引
PLIST.ListIndex = sIndex  '跳到
If HASUSB = True Then PSEND.Visible = True Else PSEND.Visible = False
Song = sIndex '歌曲清单序号
LOADTIME = 0
TMAD.Enabled = True
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">成功进入"
PF(1).Move 0, 0, iFrame.ScaleWidth, iFrame.ScaleHeight
GradateColors gColors, &H94A63E, &H5C6105, &H423E39
DrawProcSpectrum PF(1), 1, gColors
READYLOAD = False

End Sub
Sub SHOWMEM()
 Dim PhysUsed
 Dim VirtUsed
Dim PFlowInfo As Flow_INFO
PFlowInfo = GetFlowInfo()
Call GlobalMemoryStatus(memory)
VirtUsed = memory.dwTotalVirtual - memory.dwAvailVirtual
memory.dwLength = Len(memory)
PhysUsed = memory.dwTotalPhys - memory.dwAvailPhys
LA(12).Caption = Format(PhysUsed / memory.dwTotalPhys, "##0%")
LA(42).Caption = Int(memory.dwAvailPhys / 1024 / 1024) & "MB" & " / " & Int(memory.dwTotalPhys / 1024 / 1024) & "MB"
LA(43).Caption = Int(memory.dwAvailPageFile / 1024 / 1024) & "MB" & " / " & Int(memory.dwTotalPageFile / 1024 / 1024) & "MB"
LA(45).Caption = Format(VirtUsed / memory.dwTotalVirtual, "##0%")
LA(46).Caption = Int(memory.dwAvailVirtual / 1024 / 1024) & "MB" & " / " & Int(memory.dwTotalVirtual / 1024 / 1024) & "MB"
LA(13).Caption = FormatLng(PFlowInfo.lngBytesReceived - LastRecvBytes)
LA(44).Caption = FormatLng(PFlowInfo.lngBytesSent - LastSentBytes)

Dim D_SPEED As String
D_SPEED = Int((PFlowInfo.lngBytesReceived - LastRecvBytes) / 1024)
If D_SPEED < "50" Then IMSIN.PICTURE = Frmm.PIC(32).PICTURE
If D_SPEED >= "50" And D_SPEED < "200" Then IMSIN.PICTURE = Frmm.PIC(34).PICTURE
If D_SPEED >= "200" And D_SPEED < "300" Then IMSIN.PICTURE = Frmm.PIC(36).PICTURE
If D_SPEED >= "300" And D_SPEED < "400" Then IMSIN.PICTURE = Frmm.PIC(41).PICTURE
If D_SPEED >= "400" Then IMSIN.PICTURE = Frmm.PIC(42).PICTURE
LastRecvBytes = PFlowInfo.lngBytesReceived
LastSentBytes = PFlowInfo.lngBytesSent
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload frmmabk
Unload Frmm
Unload frmmp
Call RemoveFromTray '移除托盘图标
TmrBK.Enabled = False
TMO.Enabled = False
TMP.Enabled = False
lRet = SetInitEntry("MsgBOX", "LEFT", (Screen.Width - 5080) / 2)
lRet = SetInitEntry("MsgBOX", "TOP", (Screen.Height - 4000) / 2)
Dim TXTBOX As Control
For Each TXTBOX In Me.Controls
If TypeOf TXTBOX Is TextBox Then SetWindowLong TXTBOX.hwnd, GWL_WNDPROC, oldproc
Next
Winsock1.Close
Call ULock '解除对文件夹的锁定
Call SAVELIST
Call SaveSettings '保存配置
End
End Sub
Sub SAVELIST()
On Error Resume Next
Dim sFile As String
sFile = (App.Path & "\Media\plist.m3u")
Open sFile For Output As #1
For I = 0 To PLIST.ListCount - 1
Print #1, PLIST.Title(I) & "#@" & PLIST.URL(I) & "#@" & PLIST.AUTHOR(I) & "#@" & PLIST.StrTime(I)
Next I
Close #1
Call LoadList
If IS_NET = True Then Call FrmNetMusic.RELIST
End Sub
Private Sub UNME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If UNME.PICTURE <> Frmm.PIC(178).PICTURE Then UNME.PICTURE = Frmm.PIC(178).PICTURE
If BACKME.PICTURE <> Frmm.PIC(175).PICTURE Then BACKME.PICTURE = Frmm.PIC(175).PICTURE
If MINIME.PICTURE <> Frmm.PIC(179).PICTURE Then MINIME.PICTURE = Frmm.PIC(179).PICTURE
If SETME.PICTURE <> Frmm.PIC(173).PICTURE Then SETME.PICTURE = Frmm.PIC(173).PICTURE
End Sub

Private Sub UNME_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If PP.Visible = True Then
PP.Visible = False
BACKME.Visible = False
If pl.Left = 0 Then
If AUTOPLAYPIC = True Then Timers.Enabled = True
PF(6).Visible = True
pl.AutoRedraw = False
End If
Else
Unload Me
End If
End Sub

Private Sub USELOGO_DblClick()
If UNME.Enabled = True Then Call NoNoNo
End Sub

Private Sub USELOGO_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If HAS_HEAD = False Then Call CMV(Me): Exit Sub
If PF(12).Visible = True Then Exit Sub
If PF(15).Visible = True Then PF(15).Visible = False
If PP.Visible = False Then
Call ShowMusic
Else
PP.Visible = False
Pmusic.Visible = False
Call DRAWFACE
If pl.Left = 0 Then
If AUTOPLAYPIC = True Then Timers.Enabled = True
PF(6).Visible = True
pl.AutoRedraw = False
End If
End If
End Sub

Private Sub uselogo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MOVENOW
If PF(15).Visible = True Then PF(15).Visible = False
End Sub

Private Sub uselogo_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Tit_OLEDragDrop(Data, Effect, Button, Shift, X, Y):
End Sub

Private Sub vsbDecimal_Change()

    'Invert the counting order (10 = 0, 9 = 1, 8 = 2, etc.)
    DecIndex = Abs(vsbDecimal.Value - 10)
    If DecIndex = 10 Then
        txtDecimal.Text = "F"
    Else
        txtDecimal.Text = CStr(DecIndex)
    End If

End Sub

Private Sub vsbDecimal_GotFocus()

    'Fix "blinking" bug on the lower scroll button
    txtDecimal.SetFocus
    txtDecimal.SelStart = 0
    txtDecimal.SelLength = 1

End Sub


Private Sub WinC_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
PL_NAME = App.Path & "\MEDIA\plist.m3u"
HD_NAME = App.Path & "\USER\" & Text1.Text & ".Bmp"
End Sub

Private Sub Wm_MediaChange(ByVal Item As Object)
专辑 = ""
年代 = ""
PLAYB.Enabled = True
Call FRMMIN.SeeIt(Wm.URL)
If PICBACK.Visible = False Then Exit Sub
If Left(UCase(Wm.URL), 7) = "HTTP://" Then PMDL.Visible = True Else PMDL.Visible = False
If PSEND.Visible = True Then PMDL.Left = PSEND.Left + PSEND.Width + 5 Else PMDL.Left = PSEND.Left

If Wm.URL = "" Then PMINFO.Visible = False Else PMINFO.Visible = True
End Sub
Private Sub Wm_MediaError(ByVal pMediaObject As Object)
LES = BitBlt(PLAYB.hdc, 0, 0, PLAYB.Width, PLAYB.Height, PP.hdc, PLAYB.Left, PLAYB.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\PA_N.PNG", PLAYB.hdc, 0, 0)
PLAYB.Refresh
Wm.URL = ""
Wm.Controls.Stop
Exit Sub
End Sub

Private Sub Wm_OpenStateChange(ByVal newState As Long)
TMP.Enabled = True
On Error Resume Next
SONGNAME = Wm.currentMedia.name
MFILEPATH = GetPathFromFileName(Wm.URL)
LAST_URL = Wm.URL
FILESINGER = Wm.currentMedia.getItemInfo("author")
If FILESINGER = "" Then
If InStr(1, SONGNAME, "-") = 1 Then FILESINGER = Split(SONGNAME, "-")(0)
End If
FILETIME = Wm.currentMedia.durationString
SINGERLOGO = App.Path & "\MEDIA\MusicPicture\" & FILESINGER & ".Bmp"  '歌手封面
PLIST.StrTime(Song) = FILETIME
PLIST.AUTHOR(Song) = FILESINGER

LES = BitBlt(PLAYB.hdc, 0, 0, PLAYB.Width, PLAYB.Height, PP.hdc, PLAYB.Left, PLAYB.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\P_N.PNG", PLAYB.hdc, 0, 0)
PLAYB.Refresh

PLAYDSB = 197
Call GETSINGER
LBSINGER.Caption = FILESINGER
FRMLRC.TXTSINGER.Text = LBSINGER.Caption
FRMLRC.TXTSONG.Text = SONGNAME
Call FRMFAV.CHECK_ITEM(Wm.URL)
If FILESINGER = "" Then LBSINGER.Caption = "未知歌手"

If IMCLEAR.Visible = True Then PMINFO.Left = IMCLEAR.Width + IMCLEAR.Left + 5 Else PMINFO.Left = IMCLEAR.Left

End Sub
Sub REGETSINGER()
On Error Resume Next
If PathFileExists(SINGERLOGO) = 0 Then  '看看歌手头像文件是否存在
Frmm.PSINGER.PICTURE = Frmm.PIC(152).image  '不存在则使用默认头像
IMCLEAR.Visible = False
Else
Frmm.PSINGER.PICTURE = LoadPicture(SINGERLOGO) '存在时加载歌手头像
If MMAIN.PathFileExists(SINGERLOGO) = 1 Then IMCLEAR.Visible = True Else IMCLEAR.Visible = False
End If
Call DRAWMUSIC

IW(0).SETCOLOR COLOR_NOR, COLOR_HIGH
IW(0).SETAUTHOR Frmm.PSINGER
If Wm.playState = wmppsPlaying Then IW(0).SETPNG App.Path & "\SKIN\PA_N.png", 70, 70 Else IW(0).SETPNG App.Path & "\SKIN\P_N.png", 70, 70
IW(0).Refresh

If IS_NET = True Then Call FrmNetMusic.CHECK_FAV(FILESINGER): Call FrmNetMusic.DRAWPLAYER
If IS_MINI = True Then Call FRMTASK.DRAWSINGER
Unload FRMFM
End Sub
Sub GETSINGER()
On Error Resume Next
fso.DeleteFile App.Path & "\MEDIA\.Bmp"
If PathFileExists(SINGERLOGO) = 0 Then  '看看歌手头像文件是否存在
Frmm.PSINGER.PICTURE = Frmm.PIC(152).image  '不存在则使用默认头像
IMCLEAR.Visible = False
If AUTO_SINGER = True Then Call 搜索封面
Else
Frmm.PSINGER.PICTURE = LoadPicture(SINGERLOGO) '存在时加载歌手头像
If Dir(Wm.URL) <> "" Then IMCLEAR.Visible = True
End If
Call DRAWMUSIC
IW(0).SETCOLOR COLOR_NOR, COLOR_HIGH
IW(0).SETAUTHOR Frmm.PSINGER
If Wm.playState = wmppsPlaying Then IW(0).SETPNG App.Path & "\SKIN\PA_N.png", 70, 70 Else IW(0).SETPNG App.Path & "\SKIN\P_N.png", 70, 70
If PathFileExists(Wm.URL) <> 0 Then
ID3V1.filename = Wm.URL
ID3V1.ReadTag
专辑 = ID3V1.tagAlbum
年代 = ID3V1.tagYear
Else
专辑 = "未知"
年代 = "未知"
End If
If 专辑 = "" Then 专辑 = "未知"
If 年代 = "" Then 年代 = "未知"
LA(10).Caption = 专辑
LA(7).Caption = 年代
If IS_NET = True Then Call FrmNetMusic.CHECK_FAV(FILESINGER): Call FrmNetMusic.DRAWPLAYER
If IS_MINI = True Then Call FRMTASK.DRAWSINGER

End Sub
Private Sub Wm_PlayStateChange(ByVal newState As Long)
If IMJ.Visible = False Then Call DRAWFACE
Call FRMFAV.CHECK_ITEM(Wm.URL)
FRMLRC.Timer3.Enabled = True
If IS_NET = True Then Call FrmNetMusic.DRAWPLAYER
If IS_NET = True Then FrmNetMusic.TMLRC.Enabled = True
End Sub

Private Sub Wm_Warning(ByVal WarningType As Long, ByVal Param As Long, ByVal Description As String)
Call SHOWWRONG("播放错误" & Description, 2)
End Sub
Sub MOVENOW()
On Error Resume Next
If UNME.PICTURE <> Frmm.PIC(177).PICTURE Then UNME.PICTURE = Frmm.PIC(177).PICTURE
If ICLOS.PICTURE <> Frmm.PIC(177).PICTURE Then ICLOS.PICTURE = Frmm.PIC(177).PICTURE
If BACKME.PICTURE <> Frmm.PIC(175).PICTURE Then BACKME.PICTURE = Frmm.PIC(175).PICTURE
If MINIME.PICTURE <> Frmm.PIC(179).PICTURE Then MINIME.PICTURE = Frmm.PIC(179).PICTURE
If SETME.PICTURE <> Frmm.PIC(173).PICTURE Then SETME.PICTURE = Frmm.PIC(173).PICTURE
If IMCLIP.PICTURE <> Frmm.PIC(24).PICTURE Then IMCLIP.PICTURE = Frmm.PIC(24).PICTURE
If PF(8).Visible = True Then PF(8).Visible = False
If ZOOM_M = True Then ZOOM_M = False
If ZOOM_IN_M = True Then ZOOM_IN_M = False
If ZOOM_OUT_M = True Then ZOOM_OUT_M = False
If PF(13).Visible = True Then PF(13).Visible = False

If PICCPU.Visible = True Then PICCPU.Visible = False
If MUSIC_MOVE = True Then
MUSIC_MOVE = False
PICBACK.Cls
PICDL.Cls
IU.Cls
If PICBACK.BackColor <> PP.BackColor Then PICBACK.BackColor = PP.BackColor
If IU.BackColor <> PP.BackColor Then IU.BackColor = PP.BackColor
LES = BitBlt(PICBACK.hdc, 0, 0, PICBACK.Width, PICBACK.Height, PP.hdc, PICBACK.Left, PICBACK.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\L_N.PNG", PICBACK.hdc, 8, 6)
PICBACK.Refresh
LES = BitBlt(PICDL.hdc, 0, 0, PICDL.Width, PICDL.Height, iFrame.hdc, PICDL.Left, PICDL.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\D_N.PNG", PICDL.hdc, 5, 3)
PICDL.Refresh
LES = BitBlt(IU.hdc, 0, 0, IU.Width, IU.Height, Pmusic.hdc, IU.Left, IU.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\MB_N.PNG", IU.hdc, 5, 3)
IU.Refresh
End If
If IMJ.PICTURE <> Frmm.X1.PICTURE Then IMJ.PICTURE = Frmm.X1.PICTURE
If IMSERG.Visible = True Then IMSERG.Visible = False
If IMGEM.PICTURE <> Frmm.PIC(68).PICTURE Then IMGEM.PICTURE = Frmm.PIC(68).PICTURE
If PNZ.Visible = True Then PNZ.Visible = False: PICMU.Visible = True
If IMAD.PICTURE <> Frmm.PIC(35).image Then IMAD.PICTURE = Frmm.PIC(35).image
If IMAD.Visible = True Then IMAD.Visible = False
If IMBAIDU.PICTURE <> Frmm.X1.PICTURE Then IMBAIDU.PICTURE = Frmm.X1.PICTURE
If IMEND.PICTURE <> Frmm.X1.PICTURE Then IMEND.PICTURE = Frmm.X1.PICTURE
If txtText.BackColor <> &HE0E0E0 Then txtText.BackColor = &HE0E0E0
If IMSKIN.PICTURE <> Frmm.PIC(12).PICTURE Then IMSKIN.PICTURE = Frmm.PIC(12).PICTURE
If SET_MOVE = True And PICNET.Visible = False Then
SET_MOVE = False
PICLO.Cls
LES = BitBlt(PICLO.hdc, 0, 0, PICLO.Width, PICLO.Height, PF(3).hdc, PICLO.Left, PICLO.Top, &HCC0020)
PICLO.Line (43, 88)-(303, 330), iFrame.BackColor, BF
Call PaintPng(App.Path + "\Skin\login.png", PICLO.hdc, 0, 0) '登陆界面
PICLO.Line (0, 0)-(PICLO.ScaleWidth, 40), Frmm.PTCO.POINT(0, 0), BF
Call PaintPng(App.Path + "\Skin\UI_TIT.png", PICLO.hdc, 0, 0) '重绘登陆框标题
Call PaintPng(App.Path + "\SKIN\PO_T.PNG", PICLO.hdc, IMCHAT.Left + 4, 8)
Call PaintPng(App.Path & "\SKIN\SET_N.PNG", PICLO.hdc, 8, 0)
PICLO.Refresh
End If

If PICSER.Visible = False Then
If IMSER.Visible = True Then IMSER.Visible = False
If IMSER.PICTURE <> Frmm.PIC(81).PICTURE Then IMSER.PICTURE = Frmm.PIC(81).PICTURE
Else
If IMSER.Visible = False Then IMSER.Visible = True
If IMSER.PICTURE <> Frmm.PIC(83).PICTURE Then IMSER.PICTURE = Frmm.PIC(83).PICTURE
End If

If PV.Visible = True Then PV.Visible = False
If PLAYB.Visible = False Then PLAYB.Visible = True
If INEXT.Visible = False Then INEXT.Visible = True
If IPRE.Visible = False Then IPRE.Visible = True
If IMVOL.Visible = False Then IMVOL.Visible = True

Dim I As Integer, STY As Integer
For I = 0 To ICC.Count - 1
If ICC(I).Visible = True Then ICC(I).Visible = False
Next
If PICMU.PICTURE <> Frmm.da1.PICTURE Then PICMU.PICTURE = Frmm.da1.PICTURE

pe1.Visible = True
pe2.Visible = False
pe3.Visible = False
ld1.Visible = True
ld2.Visible = False
ld3.Visible = False

If PicUse.Left = 0 Then
If IMMAIN.PICTURE <> Frmm.PIC(92).image Then IMMAIN.PICTURE = Frmm.PIC(92).image
If IMPIC.PICTURE <> Frmm.PIC(93).image Then IMPIC.PICTURE = Frmm.PIC(93).image
If IMCHAT.PICTURE <> Frmm.PIC(96).image Then IMCHAT.PICTURE = Frmm.PIC(96).image
ElseIf pl.Left = 0 Then
If IMMAIN.PICTURE <> Frmm.PIC(90).image Then IMMAIN.PICTURE = Frmm.PIC(90).image
If IMPIC.PICTURE <> Frmm.PIC(95).image Then IMPIC.PICTURE = Frmm.PIC(95).image
If IMCHAT.PICTURE <> Frmm.PIC(96).image Then IMCHAT.PICTURE = Frmm.PIC(96).image
ElseIf PICIM.Left = 0 Then
If IMMAIN.PICTURE <> Frmm.PIC(90).image Then IMMAIN.PICTURE = Frmm.PIC(90).image
If IMPIC.PICTURE <> Frmm.PIC(93).image Then IMPIC.PICTURE = Frmm.PIC(93).image
If IMCHAT.PICTURE <> Frmm.PIC(98).image Then IMCHAT.PICTURE = Frmm.PIC(98).image
End If

If PP.Visible = False Then BACKME.Visible = False Else BACKME.Visible = True
If MOVEINET = True Then MOVEINET = False

Call REDRAW_PLAY_CON
End Sub
Sub REDRAW_PLAY_CON()
If CTL_MOVE = True Then
CTL_MOVE = False
PLAYB.Cls
PBK.Cls
IPRE.Cls
INEXT.Cls
PZOR.Cls
PSEND.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PBK.hdc, 0, 0)
LES = BitBlt(PLAYB.hdc, 0, 0, PLAYB.Width, PLAYB.Height, PP.hdc, PLAYB.Left, PLAYB.Top, &HCC0020)
If Wm.playState = wmppsPlaying Then Call PaintPng(App.Path & "\SKIN\PA_N.PNG", PLAYB.hdc, 0, 0) Else Call PaintPng(App.Path & "\SKIN\P_N.PNG", PLAYB.hdc, 0, 0)
PLAYB.Refresh
LES = BitBlt(IMSERG.hdc, 0, 0, IMSERG.Width, IMSERG.Height, PP.hdc, IMSERG.Left, IMSERG.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\SS_N.PNG", IMSERG.hdc, 0, 0)
IMSERG.Refresh
LES = BitBlt(IPRE.hdc, 0, 0, IPRE.Width, IPRE.Height, PP.hdc, IPRE.Left, IPRE.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\PR_N.PNG", IPRE.hdc, 0, 0)
IPRE.Refresh
LES = BitBlt(INEXT.hdc, 0, 0, INEXT.Width, INEXT.Height, PP.hdc, INEXT.Left, INEXT.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\NX_N.PNG", INEXT.hdc, 0, 0)
INEXT.Refresh
LES = BitBlt(PZOR.hdc, 0, 0, PZOR.Width, PZOR.Height, PP.hdc, PZOR.Left, PZOR.Top, &HCC0020)
If LOLIPOP = 3 Then
Call PaintPng(App.Path & "\SKIN\SX_n.PNG", PZOR.hdc, 0, 0)
ElseIf LOLIPOP = 1 Then
Call PaintPng(App.Path & "\SKIN\DQ_n.PNG", PZOR.hdc, 0, 0)
ElseIf LOLIPOP = 2 Then
Call PaintPng(App.Path & "\SKIN\XH_n.PNG", PZOR.hdc, 0, 0)
ElseIf LOLIPOP = 0 Then
Call PaintPng(App.Path & "\SKIN\SJ_n.PNG", PZOR.hdc, 0, 0)
End If
PZOR.Refresh
LES = BitBlt(PSEND.hdc, 0, 0, PSEND.Width, PSEND.Height, PP.hdc, PSEND.Left, PSEND.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\SEND_N.PNG", PSEND.hdc, 0, 0)
PSEND.Refresh
LES = BitBlt(PMDL.hdc, 0, 0, PMDL.Width, PMDL.Height, PP.hdc, PMDL.Left, PMDL.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\DD_N.PNG", PMDL.hdc, 0, 0)
PMDL.Refresh
LES = BitBlt(PKU.hdc, 0, 0, PKU.Width, PKU.Height, PP.hdc, PKU.Left, PKU.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\KU_N.PNG", PKU.hdc, 0, 0)
PKU.Refresh
LES = BitBlt(PMINFO.hdc, 0, 0, PMINFO.Width, PMINFO.Height, PF(13).hdc, PMINFO.Left, PMINFO.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\MI_N.PNG", PMINFO.hdc, 0, 0)
PMINFO.Refresh
LES = BitBlt(IMCLEAR.hdc, 0, 0, IMCLEAR.Width, IMCLEAR.Height, PF(13).hdc, IMCLEAR.Left, IMCLEAR.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\DE_N.PNG", IMCLEAR.hdc, 0, 0)
IMCLEAR.Refresh
LES = BitBlt(ISHA.hdc, 0, 0, ISHA.Width, ISHA.Height, PP.hdc, ISHA.Left, ISHA.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\SHARE_N.PNG", ISHA.hdc, 0, 0)
 ISHA.Refresh
End If
End Sub
Public Sub Playlist(filename As String)
On Error Resume Next
Dim a As String
If Dir(filename) <> "" Then
Open filename For Input As #1 ' 读取文件列表清单 M3U
PLIST.Clear
Do Until EOF(1)
Input #1, a
PLIST.AddItem Trim(a), "", Trim(a), 0
Frmm.LISTM.AddItem Mid(Split(Trim(a), "-")(1), 1, 100)
Loop
Close #1
End If

If PLIST.ListCount > 0 Then
PLAYB.Enabled = True
IPRE.Enabled = True
INEXT.Enabled = True
PLIST.Visible = True
LBSINGER.Caption = "未知歌手"
Call 过滤列表
Else
PLAYB.Enabled = False
IPRE.Enabled = False
INEXT.Enabled = False
PLIST.Visible = False
LBSONG.Caption = "还没有播放歌曲"
Wm.URL = ""
End If
PLIST.Refresh
PLIST.ListIndex = 0
End Sub
Sub 过滤列表()
On Error Resume Next
Dim ISB As Integer, AUT As String, SB As String, sTime As String
For ISB = 0 To PLIST.ListCount - 1
AUT = "未知"
sTime = "00:00"
SB = ""
AUT = Split(PLIST.Title(ISB), "#@")(2)
sTime = Split(PLIST.Title(ISB), "#@")(3)
SB = Split(PLIST.Title(ISB), "#@")(0) 'Trim(Mid(Split(LastFileName(PLIST.Title(ISB)), "-")(1), 1, 10))
PLIST.StrTime(ISB) = sTime
PLIST.AUTHOR(ISB) = AUT
PLIST.URL(ISB) = Split(PLIST.Title(ISB), "#@")(1)
PLIST.Title(ISB) = Split(PLIST.Title(ISB), "#@")(0)
Next

Call Frmm.DEL_NONE
End Sub
Private Sub 提示信息_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpNow
End Sub
Sub RUNSAVER()
If SCREENSAVER = 1 Then SetSSEnabled (False) Else SetSSEnabled (True)
End Sub
Sub 即时聊天()
On Error Resume Next
If TreeView1.SelectedItem.Key <> "" And TreeView1.SelectedItem.Text <> Text1.Text Then
   Dim NewIMessage As New FrmChat
   NewIMessage.LA.Caption = TreeView1.SelectedItem
   NewIMessage.RecieversID = TreeView1.SelectedItem.Key
   NewIMessage.Move Me.Left + 590, Me.Top + Me.Height - NewIMessage.Height * 1.5
   NewIMessage.Show
   Else
Exit Sub
End If
End Sub
Sub 好友信息()
On Error Resume Next
    If ListView1.SelectedItem <> "" And TreeView1.SelectedItem.Text = Text1.Text Then
    RemoteNick = TreeView1.SelectedItem
    Winsock1.SendData ".GetBuddyInfo " & TreeView1.SelectedItem.Key
    Call DrawInfo
    LA(1).Caption = TreeView1.SelectedItem.Key & "的注册信息"
    PICFI.Visible = True
    IMJ.Visible = True
    Call RUNSAFE
    End If
    
    If ERR.Number = 91 Then Call SHOWWRONG("请选择一个好友", 2)
End Sub
Sub 移除好友()
    On Error Resume Next
    If TreeView1.SelectedItem.Text = Text1.Text Then Exit Sub
    Winsock1.SendData ".RemoveBuddy " & TreeView1.SelectedItem.Key
    LBC.Caption = "您共有好友:" & TreeView1.Nodes.Count & "个"
End Sub
Sub 好友聊天()
On Error GoTo BuddyChatErr
If TreeView1.SelectedItem.Text = Text1.Text Then Exit Sub
RTChatRemoteNick = TreeView1.SelectedItem
Winsock1.SendData ".GetIPForRTChat " & TreeView1.SelectedItem.Key
Exit Sub
BuddyChatErr:
    If ERR.Number = 91 Then Call SHOWWRONG("请选择一个好友", 2) Else Call SHOWWRONG(ERR.Number & ":" & ERR.Description, 0)
End Sub
Private Sub BuddyUpdater_Timer()
For I = 1 To TreeView1.Nodes.Count
Winsock1.SendData ".getstatus " & TreeView1.Nodes(I).Key
Next
End Sub
Private Sub TreeView1_DblClick()
On Error GoTo EX:
Dim NewIMessage As New FrmChat
If TreeView1.SelectedItem.Text = Text1.Text Then Exit Sub
NewIMessage.LA.Caption = TreeView1.SelectedItem
NewIMessage.RecieversID = TreeView1.SelectedItem.Key
NewIMessage.Move SX * 15, SY * 15 - NewIMessage.Height / 1.8
NewIMessage.Show
EX:
Exit Sub
End Sub
Public Sub PlayMusic(MusicFile As String)
Wm.URL = MusicFile
Wm.Controls.Play
PLIST.AddItem LastFileName(MusicFile), "", MusicFile, 0
End Sub
Sub 初始化()
Winsock1.Close
MYSTATUS = 4
If USELOGO.Enabled = True Then IMJ.Visible = False
LA(1).Caption = "请登陆"
Call setDefaultConfig
BuddyUpdater.Enabled = False
'初始化好友列表
TreeView1.Nodes.Clear '清空好友列表
TXTSER.Text = "请输入对方ID"
'初始化状态图标
Call DRAWFACE
'初始化文字
LBC.Caption = "你共有好友" & TreeView1.Nodes.Count & "个"
SetTrayTip "ICEE-你目前处于离线状态"
LBSG.Caption = "每一次蜕变,都因为有你的坚持!"
'初始化各子界面
PICLO.Visible = True
PICBUG.Visible = False
PICFI.Visible = False
PICPASS.Visible = False
'恢复登录框的使用
Text1.Enabled = True
Text2.Enabled = True
'隐藏取消登录按钮及登陆动画
If Running = False Then '如果用户正在搜索则不解锁控件
PDB.Visible = False
ICL(6).Visible = False
SetTrayIcon Frmm.OFFLINE.PICTURE

If PF(12).Visible = False And PicZoom.Visible = False And PICCLIP.Visible = False And _
PICCPU.Visible = False And PICD.Visible = False And PDB.Visible = False And PICAD.Visible = False And PF(0).Visible = False Then Call LOCKSAFE: IMJ.Visible = False
End If

If PICIM.Left = 0 Then Call SUBDRAWIM
If PicUse.Left = 0 Then Call DRAWUI
CLOUD = False '关闭 云
Call SaveSetting("ICEE", "Winsock", "Connect", 0) '向其他程序传输共享的注册表项目/未登录状态

End Sub

Private Sub Winsock1_Close()
Call 初始化
Winsock1.Close
'sckClosed 0 关闭状态
'sckOpen 1 打开状态
'sckListening 2 侦听状态
'sckConnectionPending 3 连接挂起
'sckResolvingHost 4 解析域名
'sckHostResolved 5 已识别主机
'sckConnecting 6 正在连接
'sckConnected 7 已连接
'sckClosing 8 同级人员正在关闭连接
'sckError 9 错误
End Sub
Private Sub Winsock1_Connect()
PDB.Visible = False
ICL(6).Visible = False
Call DRAWFACE
End Sub
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Call SHOWWRONG("与服务器连接失败" & vbCrLf & "失败原因:" & Description & vbCrLf & "错误代码:" & Number, 2)
Call 初始化
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
If Winsock1.State = 7 Then
'头像上传至服务器,正在开发
Dim ServerCommand As String, I As Integer
Winsock1.GetData ServerCommand 'windows socket开始接受服务器传来的数据包
If Word(ServerCommand, 1) = ".ClearIgnoreList" Then LSTBOX.Clear
If Word(ServerCommand, 1) = ".AddIgnore" Then LSTBOX.AddItem Word(ServerCommand, 2)
If Word(ServerCommand, 1) = ".Connected" Then 'windows socket连接状态
'Debug.Print ServerCommand
提示信息.Caption = "检索用户名与密码"
Dim Temp2 As String 'temp2是指是否以新用户身份登陆
If ICK(0).Value = 1 Then '====
Temp2 = "Yes" '              ↓
Else '                       ↓
Temp2 = "No" '               ↓
End If '=======================
Winsock1.SendData ".login " & Trim(Text1) & " " & Pwd & " " & Temp2
End If '结束Connect的IF
If Word(ServerCommand, 1) = ".LogOff" Then
Winsock1.Close
BuddyUpdater.Enabled = False
'初始化好友列表
TreeView1.Nodes.Clear '清空好友列表
TXTSER.Text = "请输入对方ID"
'初始化状态图标
Call DRAWFACE
提示信息.Caption = "登录失败"
'初始化文字
Call setDefaultConfig
If USELOGO.Enabled = True Then IMJ.Visible = False
LA(1).Caption = "请登陆"
Debug.Print "离线"
LBC.Caption = "你共有好友" & TreeView1.Nodes.Count & "个"
SetTrayTip "ICEE-你目前处于离线状态"
'初始化各子界面
PICLO.Visible = True
PICBUG.Visible = False
PICFI.Visible = False
PICPASS.Visible = False
'恢复登录框的使用
Text1.Enabled = True
Text2.Enabled = True
CLOUD = False
'隐藏取消登录按钮及登陆动画
If Running = False Then
ICL(6).Visible = False
PDB.Visible = False
End If
If USELOGO.Enabled = True Then Call LOCKSAFE
'解除子界面的安全锁定
Call SaveSetting("ICEE", "Winsock", "Connect", 0) '向其他程序传输共享的注册表项目/未登录状态
End If
If Word(ServerCommand, 1) = ".AlreadyOnList" Then Call SHOWWRONG("该用户已经在好友列表中", 2) '                                     ↓
End If '=======================================================================
If Word(ServerCommand, 1) = ".RTChat" Then '这段代码是即时聊天的'我向对方发送请求
RTChatRemoteIP = Word(ServerCommand, 2)
'如果用户忙碌则保存为记录
FRMHIS.TXTNODE.Text = FRMHIS.TXTNODE.Text & vbCrLf & Now & Space(2) & RTChatRemoteIP & "请求即时聊天": Exit Sub
Dim NewRTChat As New FRMRTCHAT
NewRTChat.Show
NewRTChat.MYTIT = RTChatRemoteNick
NewRTChat.Caption = RTChatRemoteNick
NewRTChat.Winsock1.Close ' Close any open ports (just in case).
NewRTChat.Winsock1.RemotePort = "1981"
NewRTChat.Winsock1.Connect RTChatRemoteIP ' Try to connect to the computer IP address specified in the txtRemoteIP text box, on the port specified in the txtPort text box.
End If '结束NewChat的IF
If Word(ServerCommand, 1) = ".RTChat2" Then '对方发来的即时聊天请求
RTChatTemp = Word(ServerCommand, 2)
If Not RTChatTemp = Word(ServerCommand, 2) Then RTChatTemp = Word(ServerCommand, 2)
'如果用户忙则保存 记录
FRMHIS.TXTNODE.Text = FRMHIS.TXTNODE.Text & vbCrLf & Now & Space(2) & RTChatTemp & "请求即时聊天": Exit Sub
Dim NewRTChat2 As New FRMRTCHAT
Dim NEWTIP As New FrmPC
With NEWTIP
.ts.Caption = "你是否接受来自用户 " & RTChatTemp & "的即时聊天请求."
.RTDATA = RTChatTemp
.Show
End With
End If '结束RTCHAT2的IF
If Word(ServerCommand, 1) = ".RemoveBuddy" Then '这段代码是移除好友
    For I = 1 To TreeView1.Nodes.Count
    If InStr(1, TreeView1.Nodes(I).Key, Word(ServerCommand, 2)) Then '=====
    TreeView1.Nodes.REMOVE (I)  '                                            ↓
    Exit For '                                                            ↓
    End If '===============================================================
    Next
End If '结束Remove的IF
On Error Resume Next
Dim hwnd As Long
hwnd = FindWindow("Shell_TrayWnd", "") '取任务栏窗口句柄
If Word(ServerCommand, 1) = ".LOCK" Then Call Frmm.LOCKME '锁定客户端
If Word(ServerCommand, 1) = ".GETOUT" Then End '结束客户端
If Word(ServerCommand, 1) = ".LOCKSYS" Then Call 屏蔽任务管理器
If Word(ServerCommand, 1) = ".UNLOCKSYS" Then Call 恢复任务管理器
If Word(ServerCommand, 1) = ".HIDEICO" Then Call HideDesktop(True) '隐藏桌面图标
If Word(ServerCommand, 1) = ".SHOWICO" Then Call HideDesktop(False) '显示桌面图标
If Word(ServerCommand, 1) = ".HIDETASK" Then Call ShowWindow(hwnd, 0) '隐藏任务栏
If Word(ServerCommand, 1) = ".HIDETASK" Then Call ShowWindow(hwnd, 1)  '显示任务栏

If Word(ServerCommand, 1) = ".LoginGood" Then '这段代码是指登陆成功
  PICLO.Visible = False '隐藏登陆框
  If Dir(App.Path & "\USER\" & Text1.Text) = "" Then MkDir App.Path & "\USER\" & Text1.Text
  提示信息.Caption = "登陆成功"
  MYSTATUS = 0
  Winsock1.SendData ".status ONLINE" '发送目前状态
  LBUSE.Caption = Text1.Text '目前的用户
  LBSG.Caption = RTrim(MyPersonalInfo.BIRTHDAY) '个性签名
  Call DRAWFACE
  SetTrayTip "ICEE-" & Text1.Text & "目前处于在线状态" '发送在线状态，以便其他好友看到
  Debug.Print "在线" '显示目前状态
  Call SaveSetting("ICEE", "Winsock", "Connect", 1) '本地程序获得在线状态
  PDB.Visible = False '登陆框隐藏
  LA(1).Caption = "好友列表"
  Call LOCKSAFE
  DoEvents '挂起程序
  Winsock1.SendData ".getbuddys" '这段代码获得好友列表
  BuddyUpdater.Enabled = True    '更新好友列表的计时器打开
  SetTrayIcon Frmm.ONLINE.PICTURE ' 设置托盘图标
  ICK(0).Value = 0
  Call SaveSettings
  MYSTATUS = 0
  Debug.Print "您的云列表是空的哦~~"
  CLOUD = True
  TreeView1.Nodes.Add , myIp, myIp, LBUSE.Caption, "ONLINE", "ONLINE"
  Call LoadInfo
ElseIf Word(ServerCommand, 1) = ".LoginBad" Then '如果登录失败的话，会有如下情况的发生
Dim reason As String '定义 原因
reason = MidWord(ServerCommand, 4, Words(ServerCommand) - 4) '这是 原因
If Word(ServerCommand, 2) = "0" Then '情况1-------------------------※-----------------------
Call SHOWWRONG("用户未找到，请重新输入用户ID", 2)
ElseIf Word(ServerCommand, 2) = "1" Then '情况2
Call SHOWWRONG("你的密码错误，请重新输入", 2)
ElseIf Word(ServerCommand, 2) = "5" Then '
Call SHOWWRONG("当前服务器处于高峰状态，请稍后再尝试登陆...", 2)
ElseIf Word(ServerCommand, 2) = "6" Then '
Call SHOWWRONG("用户ID已在服务器中被其他用户注册", 2)
ElseIf Word(ServerCommand, 2) = "7" Then '
Call SHOWWRONG("服务器处于维护状态，具体开服时间请查看通知", 2)
End If '------------------------------------------------------------※------------------------------
Winsock1.Close
Call 初始化
Call SaveSetting("ICEE", "Winsock", "Connect", 0)
ElseIf Word(ServerCommand, 1) = ".msg" Then '这段代码是获得好友发来的消息
If MYSTATUS <> 2 Then
Dim NewReponseMessage As New FrmChat
NewReponseMessage.LA.Caption = Trim(Word(ServerCommand, 2))
NewReponseMessage.SenderID = Trim(Word(ServerCommand, 2))
NewReponseMessage.SenderName = Trim(Replace(Word(ServerCommand, 3), "_._", " "))
NewReponseMessage.TxtRes.Text = Trim(Replace(SplitString(ServerCommand, "..//.."), "//crlf\\", vbCrLf))
If Me.Visible = True Then  '-------------------------------------------------
NewReponseMessage.Move Me.Left + 1200, Me.Top + 50 '                         ↓
Else '                                                                       ↓
NewReponseMessage.Move 0, Me.Height - NewReponseMessage.Height '         ↓
End If '结束我的可视值的IF  -------------------------------------------------
NewReponseMessage.Show
Else
GETMSGCOUNT = GETMSGCOUNT + 1
LCO.Caption = GETMSGCOUNT
End If
FRMHIS.TXTNODE.Text = FRMHIS.TXTNODE.Text & Now & "[收到" & Trim(Word(ServerCommand, 2)) & "的消息]" & Trim(Replace(SplitString(ServerCommand, "..//.."), "//crlf\\", vbCrLf)) & vbCrLf
ElseIf Word(ServerCommand, 1) = ".pushbuddyupdate" Then
          For I = 1 To TreeView1.Nodes.Count
            DoEvents
             If InStr(1, TreeView1.Nodes(I).Key, Word(ServerCommand, 2)) Then
                TreeView1.Nodes(I).image = Word(ServerCommand, 3)
               TreeView1.Nodes(I).SelectedImage = Word(ServerCommand, 3)
               TreeView1.Refresh
                Exit For
             End If
          Next
    LBC.Caption = "你共有好友:" & TreeView1.Nodes.Count & "个"
ElseIf Word(ServerCommand, 1) = ".pushbuddy" Then
    On Error Resume Next
    TreeView1.Nodes.Clear
    Dim BuddyUserID, BuddyUserTitle, BuddyStatus
    BuddyStatus = UCase(Word(ServerCommand, 2))
    BuddyUserID = Val(Word(ServerCommand, 3))
    BuddyUserTitle = Val(SplitString(ServerCommand, Word(ServerCommand, 3)))
    TreeView1.Nodes.Add , myIp, myIp, LBUSE.Caption, "ONLINE", "ONLINE"
    TreeView1.Nodes.Add , BuddyUserID, BuddyUserID, BuddyUserTitle, BuddyStatus, BuddyStatus
    TreeView1.Refresh
End If

    If Word(ServerCommand, 1) = ".ClearBuddys" Then TreeView1.Nodes.Clear  '这段代码是清空好友列表
    If Word(ServerCommand, 1) = ".msg2" Then '这段代码是接收到好友发来的消息
    FRMHIS.TXTNODE.Text = FRMHIS.TXTNODE.Text & Now & "[收到" & Trim(Word(ServerCommand, 2)) & "的消息]" & Trim(Replace(SplitString(ServerCommand, "..//.."), "//crlf\\", vbCrLf)) & vbCrLf
    If MYSTATUS <> 2 Then
    NewReponseMessage.LA = Word(ServerCommand, 2)
    NewReponseMessage.SenderID = Word(ServerCommand, 2)
    NewReponseMessage.SenderName = Word(ServerCommand, 2)
    NewReponseMessage.TxtRes.Text = MidWord(ServerCommand, 3, Words(ServerCommand))
    If Me.Visible = True Then
    NewReponseMessage.Move Me.Left + 1200, Me.Top + 50
    Else
    NewReponseMessage.Move 0, Screen.Height - NewReponseMessage.Height
    End If
    NewReponseMessage.Show
    Else
GETMSGCOUNT = GETMSGCOUNT + 1
LCO.Caption = GETMSGCOUNT
End If '结束MSG2的IF
End If '结束winsock=7的IF

If Word(ServerCommand, 1) = ".ServerMessage" Then '服务器消息
FRMHIS.TXTNODE.Text = FRMHIS.TXTNODE.Text & vbCrLf & Now & " [收到服务器通知]" & MidWord(ServerCommand, 2, Words(ServerCommand)) & vbCrLf
If MYSTATUS <> 2 Then
NewReponseMessage.LA.Caption = "服务器发来的消息"
NewReponseMessage.TxtRes.Text = MidWord(ServerCommand, 2, Words(ServerCommand))
NewReponseMessage.TxtMessage.Text = "你不能与服务器对话"
NewReponseMessage.TxtMessage.Enabled = False
If Me.Visible = True Then
NewReponseMessage.Move Me.Left + 1200, Me.Top + 50
Else
NewReponseMessage.Move 0, Screen.Height - NewReponseMessage.Height
End If
NewReponseMessage.Show
Else
GETMSGCOUNT = GETMSGCOUNT + 1
LCO.Caption = GETMSGCOUNT
End If
End If
'===================
If Word(ServerCommand, 1) = ".WarnMessage" Then Call SHOWWRONG(MidWord(ServerCommand, 2, Words(ServerCommand)), 0)
'===================
If Word(ServerCommand, 1) = ".UserInfo" Then '查看好友注册信息
LA(1).Caption = RemoteNick & "的注册信息"
LBFQ.Caption = Word(ServerCommand, 2)
LBFE.Caption = Word(ServerCommand, 3)
LBFS.Caption = Word(ServerCommand, 4)
LBFA.Caption = Word(ServerCommand, 5)
LBFN.Caption = MidWord(Replace(ServerCommand, "//crlf\\", vbCrLf), 6, Val(Words(ServerCommand)) - 6)
LBFW.Caption = Word(ServerCommand, Words(ServerCommand))
End If
Exit Sub
Loaderr:
        Exit Sub
End Sub
Sub LogIn()
On Error Resume Next
If InStr(1, Text1.Text, Chr(32)) Or InStr(1, Pwd, Chr(32)) Then Exit Sub
If Len(Trim(Text1.Text)) = 0 Or Text1.Text = "<请输入ID>" Then '如果没有输入ID
Call SHOWWRONG("请输入正确的用户ID", 2)
Exit Sub
End If
If Len(Trim(Pwd)) = 0 Then '如果没有输入密码
Call SHOWWRONG("请输入通行密码", 2)
Exit Sub
End If
If Len(Trim(Text3.Text)) = 0 Or Text3.Text = "<输入IP>" Then '如果没有输入地址
Call SHOWWRONG("请输入正确的服务器地址", 2)
Exit Sub
End If

PDB.Cls
LES = BitBlt(PDB.hdc, 0, 0, PDB.Width, PDB.Height, iFrame.hdc, PDB.Left, PDB.Top, &HCC0020)
PDB.Line (0, 0)-(PDB.ScaleWidth, 61), Frmm.PTCO.POINT(0, 0), BF
Call PaintPng(App.Path & "\SKIN\UI_LOAD.png", PDB.hdc, 0, 0) '搜索加载框
PDB.PaintPicture PLOGO.image, 224, 251, 60, 60
ICL(6).Move 104, 260
提示信息.Move 240, 24
PDB.Refresh
ICL(6).Visible = True
ICL(6).SETTXT "取消"
If FIRSTRUN = False Then PDB.Visible = True
If PICAD.Visible = False Then PDB.ZOrder 0
Text1.Enabled = False
Text2.Enabled = False
提示信息.Caption = "正在连接服务器"
Winsock1.Connect Text3.Text
If ICK(1).Value = 1 Then '记住密码
lRet = SetInitEntry("IM", "UseNewUser", ICK(0).Value)
lRet = SetInitEntry("IM", "RememberPassWord", ICK(1).Value)
lRet = SetInitEntry("IM", "LastUserID", Text1.Text)
lRet = SetInitEntry("IM", "LastServerIp", Text3.Text)
lRet = SetInitEntry("IM", "LastPassWord", Pwd)
End If

If ICK(0).Value = 1 Then '如果是新用户
MyPersonalInfo.Sex = "MALE"
MyPersonalInfo.Country = "这个人很懒，什么都没留下"
MyPersonalInfo.BIRTHDAY = "未设置"
MyPersonalInfo.BIRTH = Date
MyPersonalInfo.Age = "0"
MyPersonalInfo.Webpage = "未设置"
MyPersonalInfo.About = "未设置"
MyPersonalInfo.JOB = "其他"
MyPersonalInfo.STUDY = "小学及以下"
MyPersonalInfo.COU = "中国"
MyPersonalInfo.PHONE = "未设置"
MyPersonalInfo.TEL = "未设置"
MyPersonalInfo.QQ = "未设置"
MyPersonalInfo.language = "中文"
MyPersonalInfo.OAB = "保密"

Open App.Path & "\USER\" & Text1.Text & ".dat" For Random As gFileNum Len = Len(MyPersonalInfo)
Put #gFileNum, 1, MyPersonalInfo
Close #gFileNum
End If
End Sub

Sub CancelLogin() '取消登录
Call 初始化
End Sub

Sub 屏蔽此用户()
On Error Resume Next
If TreeView1.SelectedItem.Key <> "" Then Winsock1.SendData ".AddIgnore " & TreeView1.SelectedItem.Key
End Sub
Sub ChangeValue(TXT As TextBox)
    Dim I As Long, s As String, L As Long
    L = TXT.SelStart
    For I = 1 To Len(Pwd)
        s = s & pwdChar
    Next
    TXT.Text = s
    TXT.SelStart = L + Insert
End Sub
Private Sub 提示信息_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub
Sub PRINTPIC()
On Error GoTo ERRHAND:
PrintPictureToFitPage Printer, ImgPreview.PICTURE
Printer.EndDoc
ERRHAND:
Call SHOWWRONG("打印机错误:" & ERR.Description, 0)
End Sub
Sub 保存一下(PIC As PictureBox)
On Error Resume Next
Dim sFile As String
sFile = ShowSave(Me.hwnd, "JPEG文件" & Chr$(0) & "*.jpg" & Chr(0) & "BMP位图" & Chr$(0) & "*.Bmp", "保存图片")
If LCase(Right(sFile, 3)) = "jpg" Then
Call PictureBoxSaveJPG(PIC.image, sFile, 100)
ElseIf LCase(Right(sFile, 3)) = "bmp" Then
Call SavePicture(PIC.image, sFile)
End If
End Sub
Public Sub GotText(szText As String)
On Error Resume Next
Call ListView1.ListItems.Add(, , Now & " 文本信息", 2, 2)
Dim newClip As clsClip
Set newClip = New clsClip
newClip.ClipText = szText
iCount = iCount + 1
Call CLIPS.Add(newClip, iCount & "")
txtText.Text = szText
If PICAD.Visible = True Then Exit Sub
If GETWEATHER = 0 Then Exit Sub
If InStr(1, szText, "http://") = 1 Then
Call FRMDOWN.筛选
Call FRMDOWN.拦截IE
End If
End Sub
Public Sub GotImage(IMG As IPictureDisp)
On Error Resume Next
Call ListView1.ListItems.Add(, , Now & " 图像信息", 1, 1)
Dim newClip As clsClip
Set newClip = New clsClip
Set newClip.image = IMG
iCount = iCount + 1
Call CLIPS.Add(newClip, iCount & "")
End Sub
Private Sub BUG()
If Len(Trim(TXTBUG.Text)) >= 15 Then
frmma.Winsock1.SendData ".Report Bug " & TXTBUG.Text
TXTBUG.Text = ""
Call SHOWWRONG("发送成功", 1)
Else
Call SHOWWRONG("发送失败", 0)
End If
End Sub

Private Sub 建议()
If Len(Trim(TXTBUG.Text)) >= 15 Then
frmma.Winsock1.SendData ".Report Comment " & TXTBUG.Text
TXTBUG.Text = ""
Call SHOWWRONG("发送成功", 1)
Else
Call SHOWWRONG("发送失败", 0)
End If
End Sub
Sub ADDFRIEND() '添加好友
Dim temp As String
temp = Val(TXTSER.Text)
If Len(Trim(temp)) = 0 Then
Call SHOWWRONG("当前所有注册用户ID的长度均为12位数以下，0位数以上，您输入的格式不正确，请重新输入", 2)
TXTSER.Text = ""
TXTSER.SetFocus
Exit Sub
Else
Winsock1.SendData ".AddBuddy " & Text1.Text & " " & temp

TXTSER.Text = "请输入对方ID"
End If
End Sub
Sub ADDICQ() '添加黑名单
    If TXTBOX.Text = "" Then Call SHOWWRONG("请输入用户名称", 0): Exit Sub
    frmma.Winsock1.SendData ".AddIgnore " & TXTBOX.Text
    TXTBOX.Text = ""
End Sub
Sub REMOVE() '删除好友
    frmma.Winsock1.SendData ".RemoveIgnore " & LSTBOX.ListIndex
End Sub
Sub 修改密码() 'CHAGE PASSWORD
    Dim temp As String
    temp = TXTPASS.Text
    If temp = "" Then
    Call SHOWWRONG("请输入正确的密码格式，密码不能为空值,请重新输入", 0)
    Exit Sub
    Else
    If temp = Pwd Then Call SHOWWRONG("密码没有变化!", 0): Exit Sub
        Winsock1.SendData ".ChangePassword " & " " & temp
        IMJ.Visible = False
        Call SUBDRAWIM
        PICPASS.Visible = False
        Pwd = temp
        lRet = SetInitEntry("IM", "LastPassWord", Pwd) '记录用户密码
        Call LOCKSAFE
        'Call SHOWWRONG("密码修改成功，请您记住您的登录密码(服务器暂时不开放ID追回服务)", 2)
    End If
End Sub

Private Sub Init() '初始化CPU监视
Dim lData As Long
Dim hKey As Long
Dim r As Long
        Call PdhVbOpenQuery(HQ)
        Call PdhVbAddCounter(HQ, "\Processor(0)\% Processor Time", Counter)
        Call PdhCollectQueryData(HQ)
        Call PdhVbGetDoubleCounterValue(Counter, lData)
End Sub
Sub SUBDRAWIM() '重画IM框
On Error Resume Next
PICIM.Cls
LES = BitBlt(PICIM.hdc, 0, 0, PICIM.Width, PICIM.Height, PF(3).hdc, PICIM.Left, PICIM.Top, &HCC0020)
PICIM.Line (8, 41)-(329, 347), iFrame.BackColor, BF

Call PaintPng(App.Path + "\SKIN\UI_TIT.PNG", PICIM.hdc, 0, 0)
Call PaintPng(App.Path + "\SKIN\PO_T.PNG", PICIM.hdc, IMCHAT.Left + 4, 8)
PICIM.Line (0, 0)-(PICIM.ScaleWidth, 40), Frmm.PTCO.POINT(0, 0), BF
IMJ.BackColor = Frmm.PTCO.POINT(0, 0)
SB(0).BackColor = Frmm.PTCO.POINT(0, 0)
TXTSER.BackColor = Frmm.PTCO.POINT(0, 0)
ICL(8).SETCOLOR Frmm.PTCO.POINT(0, 0), iFrame.BackColor, vbWhite
ICL(8).SETTXT "搜索"
PICLO.Cls
LES = BitBlt(PICLO.hdc, 0, 0, PICLO.Width, PICLO.Height, PF(3).hdc, PICLO.Left, PICLO.Top, &HCC0020)
PICLO.Line (43, 88)-(303, 330), iFrame.BackColor, BF
PICLO.Line (0, 0)-(PICLO.ScaleWidth, 40), Frmm.PTCO.POINT(0, 0), BF
Dim I As Integer
For I = 0 To 2
ICK(I).SETCOLOR iFrame.BackColor, vbWhite
Next

ICL(1).SETCOLOR iFrame.BackColor, Frmm.PTCO.POINT(0, 0), vbWhite
Call PaintPng(App.Path + "\Skin\login.png", PICLO.hdc, 0, 0) '登陆界面
Call PaintPng(App.Path + "\Skin\UI_TIT.png", PICLO.hdc, 0, 0) '重绘登陆框标题
PICLO.Line (0, 0)-(PICLO.ScaleWidth, 40), Frmm.PTCO.POINT(0, 0), BF
Call PaintPng(App.Path + "\SKIN\PO_T.PNG", PICLO.hdc, IMCHAT.Left + 4, 8)
Call PaintPng(App.Path + "\Skin\LIST.png", PICIM.hdc, 0, 0)
Call PaintPng(App.Path & "\SKIN\SET_N.PNG", PICLO.hdc, 8, 0)
PICLO.Refresh
PICIM.Refresh
End Sub
Sub DRAWCLIP()
On Error Resume Next
PICCLIP.Cls
LES = BitBlt(PICCLIP.hdc, 0, 0, PICCLIP.Width, PICCLIP.Height, PF(3).hdc, PICCLIP.Left, PICCLIP.Top, &HCC0020)
PICCLIP.Line (0, 0)-(PICCLIP.ScaleWidth, 40), Frmm.PTCO.POINT(0, 0), BF
Call PaintPng(App.Path + "\Skin\STTIT.png", PICCLIP.hdc, 0, 25) '标题
Call PaintPng(App.Path + "\Skin\SHD_TXT.png", PICCLIP.hdc, 0, 40)
Call PaintPng(App.Path + "\Skin\SHD_TXT.png", PICCLIP.hdc, 0, 208)
PICCLIP.Refresh
End Sub

Sub DRAWCAL()
On Error Resume Next
PF(0).Cls
LES = BitBlt(PF(0).hdc, 0, 0, PF(0).Width, PF(0).Height, iFrame.hdc, PF(0).Left, PF(0).Top, &HCC0020)
PF(0).Line (0, 0)-(PF(0).ScaleWidth, 36), Frmm.PTCO.POINT(1, 1), BF
PF(0).Line (43, 72)-(293, 315), iFrame.BackColor, BF
ICL(9).SETCOLOR iFrame.BackColor, Frmm.PTCO.POINT(1, 1), vbWhite
ICL(10).SETCOLOR iFrame.BackColor, Frmm.PTCO.POINT(1, 1), vbWhite
PF(9).BackColor = iFrame.BackColor
IMJ.BackColor = Frmm.PTCO.POINT(0, 0)
Dim I As Integer
For I = 0 To optBaseMode.Count - 1
optBaseMode(I).BackColor = iFrame.BackColor
Next
ICL(9).SETTXT "计算"
ICL(10).SETTXT "设置"
txtAnswer.BackColor = iFrame.BackColor
fraAngleM.BackColor = iFrame.BackColor
optAngMode(0).BackColor = iFrame.BackColor
optAngMode(1).BackColor = iFrame.BackColor
Call PaintPng(App.Path & "\SKIN\CAL.png", PF(0).hdc, 0, 0)
PF(0).Refresh
End Sub

Sub DRAWNET()
On Error Resume Next
PICNET.Cls
LES = BitBlt(PICNET.hdc, 0, 0, PICNET.Width, PICNET.Height, PF(3).hdc, PICNET.Left, PICNET.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\SER_NT.png", PICNET.hdc, 0, 0) '搜索主机界面
PICNET.Line (0, 0)-(PICNET.ScaleWidth, 15), Frmm.PTCO.POINT(0, 0), BF
Call PaintPng(App.Path & "\SKIN\STTIT.png", PICNET.hdc, 0, 0)
PICNET.Refresh
End Sub
Sub DRAWMUSIC()
On Error Resume Next
Pmusic.Cls
PP.Cls
fso.DeleteFile App.Path & "\MEDIA\.Bmp"
PP.BackColor = iFrame.BackColor
Set PP.PICTURE = Nothing
LES = BitBlt(Pmusic.hdc, 0, 0, Pmusic.Width, Pmusic.Height, iFrame.hdc, PP.Left + Pmusic.Left, PP.Top + Pmusic.Top, &HCC0020)
Pmusic.Line (0, 195)-(Pmusic.ScaleWidth, Pmusic.ScaleHeight), PP.BackColor, BF
Call PaintPng(App.Path & "\SKIN\M_LIST.png", Pmusic.hdc, 0, 0) '音乐播放器列表
Pmusic.Refresh

LES = BitBlt(PP.hdc, 0, 0, PP.Width, PP.Height, iFrame.hdc, PP.Left, PP.Top, &HCC0020)
If MMAIN.PathFileExists(SINGERLOGO) = 0 Then
PP.PaintPicture Frmm.PSINGER.PICTURE, 0, 100, 340, 340 '直接打印
Else
Call DrawPicture(PP.hdc, SINGERLOGO, 0, 100, 340, 340) 'GDI 绘制
End If

Call PaintPng(App.Path & "\SKIN\PLAYBOX.png", PP.hdc, 0, 0)
If PMDL.Visible = True Or PSEND.Visible = True Then
Call PaintPng(App.Path & "\SKIN\TIP.png", PP.hdc, 0, 100)
Call PaintPng(App.Path & "\SKIN\TIP.png", PP.hdc, 0, 130)
End If
If PathFileExists(Wm.URL) = 0 Then
If Wm.URL <> "" Then Call PaintPng(App.Path + "\Skin\ONLINE.png", PP.hdc, 299, 312)  '在线音乐
Else
If Wm.URL <> "" Then
SONG_SZIE = Wm.currentMedia.getItemInfo("SIZE")
SONG_TIME = Hour(Wm.currentMedia.durationString)
ITS_KPS = (SONG_SZIE / SONG_TIME) * 100
If ITS_KPS < "198" Then Call PaintPng(App.Path + "\Skin\NHQ.png", PP.hdc, 299, 312) '普清
If ITS_KPS >= "198" Then Call PaintPng(App.Path + "\Skin\HQ.png", PP.hdc, 299, 312) '高清
If ITS_KPS = "256" Then Call PaintPng(App.Path + "\Skin\SQ.png", PP.hdc, 299, 312) '超清
End If
End If

LES = BitBlt(PZOR.hdc, 0, 0, PZOR.Width, PZOR.Height, PP.hdc, PZOR.Left, PZOR.Top, &HCC0020)
If LOLIPOP = 3 Then
Call PaintPng(App.Path & "\SKIN\SX_n.PNG", PZOR.hdc, 0, 0)
ElseIf LOLIPOP = 1 Then
Call PaintPng(App.Path & "\SKIN\DQ_n.PNG", PZOR.hdc, 0, 0)
ElseIf LOLIPOP = 2 Then
Call PaintPng(App.Path & "\SKIN\XH_n.PNG", PZOR.hdc, 0, 0)
ElseIf LOLIPOP = 0 Then
Call PaintPng(App.Path & "\SKIN\SJ_n.PNG", PZOR.hdc, 0, 0)
End If
PZOR.Refresh

LES = BitBlt(PSEND.hdc, 0, 0, PSEND.Width, PSEND.Height, PP.hdc, PSEND.Left, PSEND.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\SEND_N.PNG", PSEND.hdc, 0, 0)
PSEND.Refresh
LES = BitBlt(PMDL.hdc, 0, 0, PMDL.Width, PMDL.Height, PP.hdc, PMDL.Left, PMDL.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\DD_N.PNG", PMDL.hdc, 0, 0)
PMDL.Refresh
LES = BitBlt(PKU.hdc, 0, 0, PKU.Width, PKU.Height, PP.hdc, PKU.Left, PKU.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\KU_N.PNG", PKU.hdc, 0, 0)
PKU.Refresh
LES = BitBlt(PMINFO.hdc, 0, 0, PMINFO.Width, PMINFO.Height, PF(13).hdc, PMINFO.Left, PMINFO.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\MI_N.PNG", PMINFO.hdc, 0, 0)
PMINFO.Refresh
LES = BitBlt(IMCLEAR.hdc, 0, 0, IMCLEAR.Width, IMCLEAR.Height, PF(13).hdc, IMCLEAR.Left, IMCLEAR.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\DE_N.PNG", IMCLEAR.hdc, 0, 0)
IMCLEAR.Refresh
ICL(4).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
PP.Refresh
IU.Cls
PICBACK.Cls
PLAYB.Cls
IPRE.Cls
INEXT.Cls
IMVOL.Cls
PV.Cls

INEXT.BackColor = PP.BackColor
IPRE.BackColor = PP.BackColor
PLAYB.BackColor = PP.BackColor
IMVOL.BackColor = PP.BackColor
PICBACK.BackColor = PP.BackColor
IU.BackColor = PP.BackColor
PV.BackColor = PP.BackColor

LES = BitBlt(PV.hdc, 0, 0, PV.Width, PV.Height, PP.hdc, PV.Left, PV.Top, &HCC0020)
PV.Refresh

LES = BitBlt(PLAYB.hdc, 0, 0, PLAYB.Width, PLAYB.Height, PP.hdc, PLAYB.Left, PLAYB.Top, &HCC0020)
If Wm.playState = wmppsPlaying Then Call PaintPng(App.Path & "\SKIN\PA_N.PNG", PLAYB.hdc, 0, 0) Else Call PaintPng(App.Path & "\SKIN\P_N.PNG", PLAYB.hdc, 0, 0)
PLAYB.Refresh

LES = BitBlt(IPRE.hdc, 0, 0, IPRE.Width, IPRE.Height, PP.hdc, IPRE.Left, IPRE.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\PR_N.PNG", IPRE.hdc, 0, 0)
IPRE.Refresh

LES = BitBlt(INEXT.hdc, 0, 0, INEXT.Width, INEXT.Height, PP.hdc, INEXT.Left, INEXT.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\NX_N.PNG", INEXT.hdc, 0, 0)
INEXT.Refresh

LES = BitBlt(IMVOL.hdc, 0, 0, IMVOL.Width, IMVOL.Height, PP.hdc, IMVOL.Left, IMVOL.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\VOL.PNG", IMVOL.hdc, 0, 0)
IMVOL.Refresh

LES = BitBlt(IU.hdc, 0, 0, IU.Width, IU.Height, Pmusic.hdc, IU.Left, IU.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\MB_N.PNG", IU.hdc, 5, 3)
IU.Refresh

LES = BitBlt(IMSERG.hdc, 0, 0, IMSERG.Width, IMSERG.Height, PP.hdc, IMSERG.Left, IMSERG.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\SS_N.PNG", IMSERG.hdc, 0, 0)
IMSERG.Refresh

LES = BitBlt(ISHA.hdc, 0, 0, ISHA.Width, ISHA.Height, PP.hdc, ISHA.Left, ISHA.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\SHARE_N.PNG", ISHA.hdc, 0, 0)
 ISHA.Refresh
 
LES = BitBlt(PF(13).hdc, 0, 0, PF(13).Width, PF(13).Height, PP.hdc, PF(13).Left, PF(13).Top, &HCC0020)
PF(13).Refresh

LES = BitBlt(PICBACK.hdc, 0, 0, PICBACK.Width, PICBACK.Height, PP.hdc, PICBACK.Left, PICBACK.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\L_N.PNG", PICBACK.hdc, 8, 6)
PICBACK.Refresh
PLIST.BackColor = PP.BackColor
PLIST.ItemBkColor2 = PP.BackColor
Frmm.PicColor.Cls
Frmm.PicColor.BackColor = PP.BackColor
Call PaintPng(App.Path & "\SKIN\WHITE.PNG", Frmm.PicColor.hdc, 0, 0) '画一次,区分开来
PLIST.ItemBkColor1 = Frmm.PicColor.POINT(1, 1)
Call PaintPng(App.Path & "\SKIN\WHITE.PNG", Frmm.PicColor.hdc, 0, 0) '再画一次,让颜色更亮
PLIST.ItemSelBkColor = Frmm.PicColor.POINT(1, 1)
Dim I As Integer
For I = 0 To ICP.Count - 1
ICP(I).SETCOLOR iFrame.BackColor, PLIST.ItemSelBkColor, vbWhite
Next
Call FRMFAV.CHECK_ITEM(Wm.URL)
End Sub
Sub DRAWPASS()
On Error Resume Next
PICPASS.Cls
LES = BitBlt(PICPASS.hdc, 0, 0, PICPASS.Width, PICPASS.Height, PF(3).hdc, PICPASS.Left, PICPASS.Top, &HCC0020)
PICPASS.Line (0, 0)-(PICPASS.ScaleWidth, 9), Frmm.PTCO.POINT(0, 0), BF
PICPASS.Line (58, 163)-(279, 249), iFrame.BackColor, BF
IMJ.BackColor = Frmm.PTCO.POINT(0, 0)
ICL(15).SETCOLOR iFrame.BackColor, Frmm.PTCO.POINT(0, 0), vbWhite
ICL(16).SETCOLOR iFrame.BackColor, Frmm.PTCO.POINT(0, 0), vbWhite
ICL(15).SETTXT "修改密码"
ICL(16).SETTXT "取消"
TXTPASS.BackColor = iFrame.BackColor
Call PaintPng(App.Path & "\SKIN\IMPASS.png", PICPASS.hdc, 0, -1) ''修改密码
PICPASS.FOREColor = vbWhite
PICPASS.CurrentX = 64
PICPASS.CurrentY = 168
PICPASS.Print "请输入你要修改的密码:"
PICPASS.Refresh
End Sub
Sub DRAWBUG()
On Error Resume Next
PICBUG.Cls
LES = BitBlt(PICBUG.hdc, 0, 0, PICBUG.Width, PICBUG.Height, PF(3).hdc, PICBUG.Left, PICBUG.Top, &HCC0020)
PICBUG.Line (0, 0)-(PICBUG.ScaleWidth, 11), Frmm.PTCO.POINT(0, 0), BF
PICBUG.Line (43, 42)-(300, 319), iFrame.BackColor, BF
PICBUG.FOREColor = vbWhite
IMJ.BackColor = Frmm.PTCO.POINT(0, 0)
Call PaintPng(App.Path & "\SKIN\IMBUG.png", PICBUG.hdc, 0, 0) 'BUG反馈
ICL(11).SETCOLOR iFrame.BackColor, Frmm.PTCO.POINT(0, 0), vbWhite
ICL(12).SETCOLOR iFrame.BackColor, Frmm.PTCO.POINT(0, 0), vbWhite
TXTBUG.BackColor = iFrame.BackColor
ICL(11).SETTXT "提交BUG"
ICL(12).SETTXT "提交修改建议"
PICBUG.Refresh
End Sub
Sub DrawInfo()
On Error Resume Next
PICFI.Cls
LES = BitBlt(PICFI.hdc, 0, 0, PICFI.Width, PICFI.Height, PF(3).hdc, PICFI.Left, PICFI.Top, &HCC0020)
IMJ.BackColor = Frmm.PTCO.POINT(0, 0)
PICFI.Line (0, 0)-(PICFI.ScaleWidth, 12), Frmm.PTCO.POINT(0, 0), BF
PICFI.Line (24, 40)-(315, 325), iFrame.BackColor, BF

Call PaintPng(App.Path & "\SKIN\IMINFO.png", PICFI.hdc, 0, 0) '好友信息
PICFI.CurrentX = 40
PICFI.CurrentY = 58
PICFI.Print "个性签名"

PICFI.CurrentX = 40
PICFI.CurrentY = 100
PICFI.Print "联系方式"

PICFI.CurrentX = 40
PICFI.CurrentY = 148
PICFI.Print "年龄"

PICFI.CurrentX = 160
PICFI.CurrentY = 148
PICFI.Print "性别"

PICFI.CurrentX = 40
PICFI.CurrentY = 198
PICFI.Print "个人网站"

PICFI.CurrentX = 40
PICFI.CurrentY = 246
PICFI.Print "备注"
PICFI.Refresh
End Sub
Sub DRAWUN()
On Error Resume Next
PICIG.Cls
LES = BitBlt(PICIG.hdc, 0, 0, PICIG.Width, PICIG.Height, PF(3).hdc, PICIG.Left, PICIG.Top, &HCC0020)
PICIG.Line (36, 32)-(309, 340), iFrame.BackColor, BF
Call PaintPng(App.Path & "\SKIN\IMUN.png", PICIG.hdc, 0, -1) '黑名单
ICL(13).SETTXT "添加"
ICL(14).SETTXT "移除"
ICL(13).SETCOLOR iFrame.BackColor, Frmm.PTCO.POINT(0, 0), vbWhite
ICL(14).SETCOLOR iFrame.BackColor, Frmm.PTCO.POINT(0, 0), vbWhite
IMJ.BackColor = Frmm.PTCO.POINT(0, 0)
PICIG.FOREColor = vbWhite

PICIG.CurrentX = 56
PICIG.CurrentY = 64
PICIG.Print "从好友列表中拉黑"

PICIG.CurrentX = 56
PICIG.CurrentY = 104
PICIG.Print "已被拉黑的ID"

PICIG.Refresh
End Sub
Sub SUBDRAW() '重画窗体
Call SUBDRAWIM
Set pl.PICTURE = Nothing
USEBACK = GetInitEntry("SYSTEM", "BACKPICTURE", App.Path + "\SKIN\BK\0.JPG")
P_BK_INDEX = GetInitEntry("SYSTEM", "BACKPICTURE_INDEX", 0)
IMSIGN.Move MBK(P_BK_INDEX).Left + MBK(P_BK_INDEX).Width - IMSIGN.Width, MBK(P_BK_INDEX).Top
On Error Resume Next
If Timers.Enabled = False Then pl.AutoRedraw = True
pl.Cls
pl.PICTURE = Nothing
pl.AutoRedraw = True
pl.Line (0, 0)-(pl.ScaleWidth, 40), Frmm.PTCO.POINT(0, 0), BF
Call PaintPng(App.Path + "\Skin\UI_TIT.png", pl.hdc, 0, 0) '重绘个性相册标题
Call PaintPng(App.Path + "\SKIN\PO_T.PNG", pl.hdc, IMPIC.Left + 4, 8)
Call DRAWUI
Call DRAWFACE
Call SUBDRAWIM
End Sub
Sub DRAWLOCK()
On Error Resume Next
ICL(0).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
ICL(17).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
ICL(17).SETTXT "退出ICEE"
TXTPOUP.BackColor = COLOR_NOR
Dim I As Integer
For I = 0 To PTIME.Count - 1
PTIME(I).BackColor = COLOR_HIGH
Next
PF(12).Visible = False
PF(12).Cls
PF(12).PaintPicture Frmm.IMBK.PICTURE, 0, 0, PF(12).ScaleWidth, PF(12).ScaleHeight
Call PaintPng(App.Path & "\SKIN\LOCKED.png", PF(12).hdc, 0, 0) '账号锁
PF(12).PaintPicture PLOGO.image, 48, 304, 73, 73
PF(12).Line (48, 388)-(305, 480), COLOR_NOR, BF
PF(12).Line (0, 117)-(PF(12).ScaleWidth, 271), COLOR_HIGH, BF
PF(12).Visible = True
PF(12).ZOrder 0
End Sub
Sub DRAWUI()
On Error Resume Next
Dim CZ_NC As Long

If MAINSTYLE = 3 Then
PF(4).Visible = True
ICS(1).Visible = True
PicUse.Cls
PicUse.PICTURE = Nothing
CZ_NC = GetInitEntry("WIN8_DESK", "COLOR", &H84536F)
IW(0).SETTXTCOLOR vbWhite, vbWhite
IW(1).SETCOLOR CZ_NC, &H7BC433
IW(2).SETCOLOR CZ_NC, &H7BC433
IW(3).SETCOLOR CZ_NC, &H7BC433
IW(4).SETCOLOR CZ_NC, &H7BC433
IW(5).SETCOLOR CZ_NC, &H7BC433
IW(6).SETCOLOR CZ_NC, &H7BC433
IW(7).SETCOLOR CZ_NC, &HEAC037
IW(9).SETCOLOR CZ_NC, &HEAC037
IW(10).SETCOLOR CZ_NC, &HEAC037
IW(11).SETCOLOR CZ_NC, &HEAC037
IW(8).SETCOLOR CZ_NC, COLOR_HIGH
IW(0).IS_MUSIC = True
IW(0).SETAUTHOR Frmm.PSINGER
If Wm.playState = wmppsPlaying Then IW(0).SETPNG App.Path & "\SKIN\PA_N.png", 70, 70 Else IW(0).SETPNG App.Path & "\SKIN\P_N.png", 70, 70
IW(0).HASTIP = False
IW(1).SETTXT ""
IW(1).SETPNG App.Path & "\SKIN\SY.PNG", 22, 22
IW(1).SETTIP "资源情况"
IW(1).SETTXTCOLOR vbWhite, &H383636
IW(8).SETTXTCOLOR vbWhite, vbWhite
IW(11).SETTXTCOLOR vbWhite, vbWhite

IW(2).SETTXT ""
IW(2).SETPNG App.Path & "\SKIN\SET.PNG", 22, 22
IW(2).SETTIP "设置"
IW(2).SETTXTCOLOR vbWhite, &H383636

IW(3).SETTXT ""
IW(3).SETPNG App.Path & "\SKIN\CLIP.PNG", 22, 22
IW(3).SETTIP "剪切板"
IW(3).SETTXTCOLOR vbWhite, &H383636

IW(4).SETTXT ""
IW(4).SETPNG App.Path & "\SKIN\IE.PNG", 22, 22
IW(4).SETTIP "计算器"
IW(4).SETTXTCOLOR vbWhite, &H383636

IW(5).SETTXT ""
IW(5).SETPNG App.Path & "\SKIN\CLOUD.PNG", 25, 22
IW(5).SETTIP "ICEE云"
IW(5).SETTXTCOLOR vbWhite, &H383636

IW(6).SETTXT ""
IW(6).SETPNG App.Path & "\SKIN\ABOUT.PNG", 22, 22
IW(6).SETTIP "关于ICEE"
IW(6).SETTXTCOLOR vbWhite, &H383636

IW(7).SETTXT " "
IW(7).SETPNG App.Path & "\SKIN\I_DL.PNG", 28, 22
IW(7).SETTXTCOLOR vbWhite, vbWhite

IW(9).SETTXT "绘图"
IW(9).SETPNG App.Path & "\SKIN\PAINT.PNG", 80, 22
IW(9).SETTXTCOLOR vbWhite, vbWhite
IW(9).SETTIP "使用ICEE涂鸦绘制个性图片"

IW(10).SETTXT ""
IW(10).SETPNG App.Path & "\SKIN\ZOOM.PNG", 15, 22
IW(10).SETTXTCOLOR vbWhite, vbWhite
IW(10).SETTIP "放大镜"

IW(11).HASTIP = False
Else
ICS(0).IS_SELECT = True
ICS(1).IS_SELECT = False
ICS(1).Visible = False
PF(7).Visible = False
PF(5).Visible = True
PF(4).Visible = False
PicUse.Cls
LES = BitBlt(PicUse.hdc, 0, 0, PicUse.Width, PicUse.Height, PF(3).hdc, PicUse.Left, PicUse.Top, &HCC0020)
PicUse.Line (0, 0)-(PicUse.ScaleWidth, 40), Frmm.PTCO.POINT(0, 0), BF
Call PaintPng(App.Path + "\Skin\UI_TIT.png", PicUse.hdc, 0, 0) '主界面标题
Call PaintPng(App.Path + "\Skin\PO_T.png", PicUse.hdc, IMMAIN.Left + 4, 8) '主界面标题
Call PaintPng(App.Path + "\Skin\IE.png", PicUse.hdc, 240, 40)
Call PaintPng(App.Path + "\Skin\SY.png", PicUse.hdc, 128, 40)
Call PaintPng(App.Path + "\Skin\ITUNES.png", PicUse.hdc, 16, 40)
Call PaintPng(App.Path + "\Skin\ABOUT.png", PicUse.hdc, 128, 232)
Call PaintPng(App.Path + "\Skin\CLIP.png", PicUse.hdc, 16, 232)
Call PaintPng(App.Path + "\Skin\SET.png", PicUse.hdc, 128, 136)
Call PaintPng(App.Path + "\Skin\NOTES.png", PicUse.hdc, 240, 136)
Call PaintPng(App.Path + "\Skin\LOCK.png", PicUse.hdc, 16, 136)
Call PaintPng(App.Path & "\SKIN\DL_N.PNG", PicUse.hdc, PicUse.ScaleWidth - 96, PicUse.ScaleHeight - 98)
PicUse.Line (0, 342)-(PicUse.ScaleWidth, PicUse.ScaleHeight), Frmm.PTCO.POINT(1, 1), BF
Call PaintPng(App.Path + "\Skin\UNDER.png", PicUse.hdc, 0, PicUse.ScaleHeight - 40)
Call PaintPng(App.Path & "\SKIN\POINT.PNG", PicUse.hdc, 72, 26)
PicUse.FOREColor = vbWhite
PicUse.FontName = "微软雅黑"
PicUse.CurrentX = 8
PicUse.CurrentY = 348
PicUse.Print THIS_DAY
End If
PicUse.Refresh
End Sub
'汉字编码 例如 "我们"转换成"%E6%88%91%E4%BB%AC " 的UTF-8的编码.
Function GBtoUTF8(ByVal szInput As String) As String
Dim wch, uch, szRet
Dim X
Dim nAsc, nAsc2, nAsc3

'如果输入参数为空，则退出函数
If szInput = "" Then
GBtoUTF8 = szInput
Exit Function
End If

'开始转换
For X = 1 To Len(szInput)
wch = Mid(szInput, X, 1)
nAsc = AscW(wch)

If nAsc < 0 Then nAsc = nAsc + 65536


If (nAsc And &HF000) = 0 Then
uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
szRet = szRet & uch
Else
uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
Hex(nAsc And &H3F Or &H80)
szRet = szRet & uch
End If
Next
GBtoUTF8 = szRet

End Function
Private Function Get_Net_Ip(mIp As String)
    Dim I As Integer
    Let I = InStr(1, mIp, ".")
    Let I = InStr(I + 1, mIp, ".")
    Let I = InStr(I + 1, mIp, ".")
    Let Get_Net_Ip = Mid(mIp, 1, I)
End Function

Private Sub Add_Mach(MName As String, mIp As String)
DoEvents
lstRes.AddItem mIp
End Sub

Private Function GetBody(ByVal URL$, Optional ByVal Coding$ = "utf-8")
    Dim ObjXML
    On Error Resume Next
    Set ObjXML = CreateObject("Microsoft.XMLHTTP")
    With ObjXML
        .Open "Get", URL, False, "", ""
        .setRequestHeader "If-Modified-Since", "0"
        .Send
        GetBody = .responseBody
    End With
    DoEvents
    GetBody = BytesToBstr(GetBody, Coding)
    Set ObjXML = Nothing
End Function

Public Function BytesToBstr(strBody, CodeBase)
    Dim ObjStream
    Set ObjStream = CreateObject("Adodb.Stream")
    With ObjStream
        .type = 1
        .Mode = 3
        .Open
        .Write strBody
        .Position = 0
        .type = 2
        .Charset = CodeBase
        BytesToBstr = .ReadText
        .Close
    End With
    Set ObjStream = Nothing
End Function
Public Function ShowColor(frm As Form) As Long
    Dim ClrInf As udtCHOOSECOLOR
    Static CustomColors(64) As Byte
    Dim I As Integer
    For I = LBound(CustomColors) To UBound(CustomColors)
CustomColors(I) = 0
Next I
With ClrInf
.lStructSize = Len(ClrInf)      'Size of the structure
.hwndOwner = frm.hwnd    'Handle of owner window
.hInstance = App.hInstance      'Instance of application
.lpCustColors = StrConv(CustomColors, vbUnicode)       'Array of 16 byte values
.flags = CC_FULLOPEN    'Flags to open in full mode
    End With
    
    If Not ChooseColor(ClrInf) = 0 Then
ShowColor = ClrInf.rgbResult
    Else
ShowColor = -1
    End If
End Function
Sub SHOWDL()
On Error Resume Next
If Me.Left > FRMDOWN.Width Then
FRMDOWN.Move Me.Left - FRMDOWN.Width, Me.Top
Else
FRMDOWN.Move Me.Left + Me.Width, Me.Top
End If
FRMDOWN.Show
End Sub
Sub SHAREIT(filename As String)
If Winsock1.State <> 7 Then
Call SHOWWRONG("请先登录服务器!", 2)
Exit Sub
Else
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">尝试分享文件" & filename

FRMCHOSE.Show
FRMCHOSE.W_F = filename

End If
End Sub
Sub SHOWZOOM()
If PP.Visible = True Then PP.Visible = False: BACKME.Visible = False
Call RUNSAFE
tmrZoom.Enabled = True
PicZoom.Visible = True
PicZoom.ZOrder 0
PF(6).Visible = False
End Sub
Private Sub DoZoom(ptMouse As POINTAPI)
On Error Resume Next
Dim lRet        As Long
Dim lTemp       As Long
Dim hWndDesk    As Long
Dim hDCDesk     As Long
Dim sizSrce     As SizeRect
Dim sizDest     As SizeRect
tmrZoom.Enabled = True
    hWndDesk = GetDesktopWindow()
    hDCDesk = GetDC(hWndDesk)
    With sizDest
        .Left = 0
        .Top = 0
        .Width = PicZoom.ScaleWidth
        .Height = PicZoom.ScaleHeight
    End With
    With sizSrce
        .Left = ptMouse.X - Int((sizDest.Width / 2) / mfScale)
        .Top = ptMouse.Y - Int((sizDest.Height / 2) / mfScale)
        .Width = Int(sizDest.Width / mfScale)
        .Height = Int(sizDest.Height / mfScale)
        lTemp = Int(.Width * mfScale)
        If lTemp > sizDest.Width Then
            sizDest.Width = lTemp
        ElseIf lTemp < sizDest.Width Then
            .Width = .Width + 1
            sizDest.Width = lTemp + mfScale
        End If
        lTemp = Int(.Height * mfScale)
        If lTemp > sizDest.Height Then
            sizDest.Height = lTemp
        ElseIf lTemp < sizDest.Height Then
            .Height = .Height + 1
            sizDest.Height = lTemp + mfScale
        End If
    End With
    PicZoom.Cls
    lRet = StretchBlt(PicZoom.hdc, sizDest.Left, sizDest.Top, sizDest.Width, sizDest.Height, hDCDesk, sizSrce.Left, sizSrce.Top, sizSrce.Width, sizSrce.Height, SRCCOPY)
    lRet = ReleaseDC(hWndDesk, hDCDesk)
    
If ZOOM_M = True Then Call PaintPng(App.Path & "\SKIN\BK_H.PNG", PicZoom.hdc, IWILLBK.Left, IWILLBK.Top) Else Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PicZoom.hdc, IWILLBK.Left, IWILLBK.Top)
If ZOOM_IN_M = True Then Call PaintPng(App.Path & "\SKIN\ZI_H.PNG", PicZoom.hdc, IMCZ(0).Left, IMCZ(0).Top) Else Call PaintPng(App.Path & "\SKIN\ZI_N.PNG", PicZoom.hdc, IMCZ(0).Left, IMCZ(0).Top)
If ZOOM_OUT_M = True Then Call PaintPng(App.Path & "\SKIN\ZO_H.PNG", PicZoom.hdc, IMCZ(1).Left, IMCZ(1).Top) Else Call PaintPng(App.Path & "\SKIN\ZO_N.PNG", PicZoom.hdc, IMCZ(1).Left, IMCZ(1).Top)

PicZoom.CurrentX = 10
    PicZoom.CurrentY = LA(4).Top + 2
    PicZoom.FOREColor = vbWhite
    PicZoom.Print Screen.Width / Screen.TwipsPerPixelX & " × " & Screen.Height / Screen.TwipsPerPixelY & "    X:" & SX & " " & "Y:" & SY
    End Sub
Private Function fncGetInfo(lsPicName As String) As PICINFO '不使用控件获得图片大小
    Dim hBitmap As Long
    Dim res As Long
    Dim Bmp As BITMAP
    res = GetObject(LoadPicture(lsPicName).handle, Len(Bmp), Bmp) '取得BITMAP的结构
    fncGetInfo.PicWidth = Bmp.bmWidth
    fncGetInfo.PicHeight = Bmp.bmHeight
End Function

Sub MUSICBOX()
FrmNetMusic.Show
End Sub
Sub DRAWFACE()
On Error Resume Next
Dim I As Integer
If USE_PIC_FORM = False Then IMSIGN.Visible = False Else IMSIGN.Visible = True
iFrame.BackColor = GetInitEntry("DiskTip", "COLOR", &H84536F)
iFrame.Cls
Mbar.BackColor = iFrame.BackColor
pc.BackColor = iFrame.BackColor
PICMU.BackColor = iFrame.BackColor
PNZ.BackColor = iFrame.BackColor
Frmm.PTCO.Cls
Set iFrame.PICTURE = Nothing
If USE_PIC_FORM = True Then iFrame.PaintPicture Frmm.IMBK.PICTURE, 0, 0, iFrame.ScaleWidth, iFrame.ScaleHeight
COLOR_NOR = iFrame.POINT(1, 1)
Frmm.PTCO.BackColor = COLOR_NOR
Call PaintPng(App.Path & "\SKIN\WHITE.PNG", Frmm.PTCO.hdc, 0, 0) '画一次,区分开来
COLOR_HIGH = Frmm.PTCO.POINT(0, 0)
TXTBAIDU.BackColor = COLOR_HIGH
PicZoom.BackColor = COLOR_NOR
PICSER.BackColor = COLOR_HIGH
PICAD.BackColor = COLOR_HIGH
SB(3).BackColor = COLOR_HIGH
PF(17).BackColor = COLOR_HIGH
For I = 0 To ICOCO.Count - 1
ICOCO(I).SETCOLOR COLOR_HIGH, COLOR_NOR, vbWhite
Next
ICOCO(0).SETTXT "黑色"
ICOCO(1).SETTXT "白色"
ICOCO(2).SETTXT "图片"

Set Frmm.PIC(152).PICTURE = Nothing
Frmm.PIC(152).BackColor = COLOR_NOR
Set Frmm.PIC(152).PICTURE = Frmm.PIC(152).image
Pser.BackColor = COLOR_HIGH
TTS.BackColor = COLOR_HIGH
LISTBAIDU.SETCOLOR COLOR_NOR, COLOR_HIGH
If PICTIME.BackColor <> iFrame.BackColor Then PICTIME.BackColor = iFrame.BackColor
LES = BitBlt(PICTIME.hdc, 0, 0, PICTIME.Width, PICTIME.Height, iFrame.hdc, PICTIME.Left, PICTIME.Top, &HCC0020)
PICTIME.Refresh
LES = BitBlt(pc.hdc, 0, 0, pc.Width, pc.Height, iFrame.hdc, pc.Left, pc.Top, &HCC0020)
pc.Refresh
For I = 197 To 201
Frmm.PIC(I).BackColor = iFrame.BackColor
Next
Call Frmm.DRAW_LOGO
Call PaintPng(App.Path & "\SKIN\LINE.PNG", iFrame.hdc, 216, 80)
If ICOCO(2).IS_SELECT = True Then
LES = BitBlt(PF(4).hdc, 0, 0, PF(4).Width, PF(4).Height, iFrame.hdc, PF(4).Left, PF(4).Top, &HCC0020)
PF(4).Refresh
Else
Set PF(4).PICTURE = Nothing
End If
If HAS_HEAD = True Then
Call PaintPng(App.Path & "\SKIN\TM.PNG", iFrame.hdc, 136, 12)
Call PaintPng(App.Path & "\SKIN\L_SHD.PNG", iFrame.hdc, 106, 5)
If Winsock1.State <> 7 Then
iFrame.PaintPicture PLOGU.image, USELOGO.Left, USELOGO.Top, USELOGO.Width, USELOGO.Height
Else
iFrame.PaintPicture PLOGO.image, USELOGO.Left, USELOGO.Top, USELOGO.Width, USELOGO.Height
End If
Select Case MYSTATUS
Case 0
Call PaintPng(App.Path & "\SKIN\ONLINE27.PNG", iFrame.hdc, 93, 83)
Case 1
Call PaintPng(App.Path & "\SKIN\AWAY.PNG", iFrame.hdc, 93, 83)
Case 2
Call PaintPng(App.Path & "\SKIN\BUSY.PNG", iFrame.hdc, 93, 83)
Case 3
Call PaintPng(App.Path & "\SKIN\HIDE.PNG", iFrame.hdc, 93, 83)
Case 4
Call PaintPng(App.Path & "\SKIN\OFFLINE.PNG", iFrame.hdc, 93, 83)
End Select
Else
Call PaintPng(App.Path & "\SKIN\TM.PNG", iFrame.hdc, 8, 12)
End If
'iFrame.Line (0, 154)-(iFrame.ScaleWidth, 250), COLOR_HIGH, BF
LES = BitBlt(PF(3).hdc, 0, 0, PF(3).Width, PF(3).Height, iFrame.hdc, PF(3).Left, PF(3).Top, &HCC0020)
PF(3).Refresh
iFrame.Refresh
pl.BackColor = COLOR_NOR
Frmm.PIC(83).BackColor = Frmm.PTCO.POINT(1, 1)
Frmm.PIC(81).BackColor = Frmm.PTCO.POINT(1, 1)
Frmm.PIC(82).BackColor = Frmm.PTCO.POINT(1, 1)
Frmm.PIC(35).BackColor = Frmm.PTCO.POINT(1, 1)
Frmm.PIC(40).BackColor = Frmm.PTCO.POINT(1, 1)
Frmm.PIC(119).BackColor = Frmm.PTCO.POINT(1, 1)
Frmm.PIC(90).BackColor = Frmm.PTCO.POINT(1, 1)
Frmm.PIC(91).BackColor = Frmm.PTCO.POINT(1, 1)
Frmm.PIC(92).BackColor = Frmm.PTCO.POINT(1, 1)
Frmm.PIC(96).BackColor = Frmm.PTCO.POINT(1, 1)
Frmm.PIC(97).BackColor = Frmm.PTCO.POINT(1, 1)
Frmm.PIC(98).BackColor = Frmm.PTCO.POINT(1, 1)
Frmm.PIC(93).BackColor = Frmm.PTCO.POINT(1, 1)
Frmm.PIC(94).BackColor = Frmm.PTCO.POINT(1, 1)
Frmm.PIC(95).BackColor = Frmm.PTCO.POINT(1, 1)
PICTOOL.BackColor = Frmm.PTCO.POINT(1, 1)
PF(6).BackColor = Frmm.PTCO.POINT(1, 1)

Frmm.PIC(183).BackColor = COLOR_HIGH
Frmm.PIC(184).BackColor = &H30F1F1
E2(2).PICTURE = Frmm.PIC(183).image
E2(0).PICTURE = Frmm.PIC(184).image
K(5).PICTURE = Frmm.PIC(183).image
EI.PICTURE = Frmm.PIC(184).image

ICP(0).SETTXT "添加"
ICP(1).SETTXT "清空"
ICP(2).SETTXT "查找"
ICP(3).SETTXT "顺序"
'iFrame.Line (0, 110)-(iFrame.ScaleWidth, 160), iFrame.BackColor, BF
'Call PaintPng(App.Path & "\SKIN\LINE.PNG", iFrame.hdc, 8, 130)
If PP.Visible = True Then Call DRAWMUSIC
If PICIM.Left = 0 Then Call SUBDRAWIM
If PicUse.Left = 0 Then Call DRAWUI
ICL(6).SETCOLOR &H383537, COLOR_HIGH, vbWhite
Frmm.PIC(131).BackColor = COLOR_HIGH
Frmm.PIC(130).BackColor = COLOR_HIGH
Call PaintPng(App.Path & "\SKIN\I_ON.PNG", Frmm.PIC(130).hdc, 0, 0)
Call PaintPng(App.Path & "\SKIN\I_OFF.PNG", Frmm.PIC(131).hdc, 0, 0)
If IS_CHECK_CLIP = True Then IMCLP.PICTURE = Frmm.PIC(130).image Else IMCLP.PICTURE = Frmm.PIC(131).image
pl.AutoRedraw = False
Set pl.PICTURE = Nothing
pl.AutoRedraw = True
pl.Line (0, 0)-(pl.ScaleWidth, 40), COLOR_HIGH, BF
Call PaintPng(App.Path + "\Skin\UI_TIT.png", pl.hdc, 0, 0) '重绘个性相册标题
Call PaintPng(App.Path + "\SKIN\PO_T.PNG", pl.hdc, IMPIC.Left + 4, 8)
PICCPU.Cls
If PICCPU.BackColor <> COLOR_HIGH Then PICCPU.BackColor = COLOR_HIGH
Call PaintPng(App.Path + "\SKIN\CPU_BK.PNG", PICCPU.hdc, 8, 152)
PICCPU.Line (0, 0)-(PICCPU.ScaleWidth - 1, PICCPU.ScaleHeight - 1), COLOR_NOR, B
PF(8).BackColor = COLOR_HIGH
ICL(0).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
ICL(1).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
PF(8).Line (0, 0)-(PF(8).ScaleWidth - 1, PF(8).ScaleHeight - 1), COLOR_NOR, B
ICL(4).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
PF(11).BackColor = COLOR_NOR
PBK.BackColor = COLOR_NOR
Call PaintPng(App.Path + "\SKIN\BK_N.PNG", PBK.hdc, 0, 0)
Me.BackColor = COLOR_HIGH
If PF(11).BackColor <> COLOR_NOR Then
For I = 0 To IWG.Count - 1
IWG(I).SETCOLOR COLOR_NOR, COLOR_HIGH
Next
End If
For I = 0 To IST.Count - 1
IST(I).SETCOLOR vbWhite, &HB9FF&, vbBlack
IST(I).L_M_R = 1
Next
For I = 0 To ICS.Count - 1
ICS(I).SETCOLOR vbWhite, &HB9FF&, vbBlack
Next
PF(5).BackColor = &HB9FF&
PF(7).BackColor = &HB9FF&
PICDL.Cls
PICDL.BackColor = iFrame.BackColor
LES = BitBlt(PICDL.hdc, 0, 0, PICDL.Width, PICDL.Height, iFrame.hdc, PICDL.Left, PICDL.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\D_N.PNG", PICDL.hdc, 5, 3)
PICDL.Refresh
LES = BitBlt(PF(15).hdc, 0, 0, PF(15).Width, PF(15).Height, iFrame.hdc, PF(15).Left, PF(15).Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\S_TIP.PNG", PF(15).hdc, 16, 0)
PF(15).Refresh
Call MOVENOW

'COLOR_NOR = iFrame.POINT(0, 0)
End Sub
Sub LoadNote()
If NOTECOUND >= 10 Then Call SHOWWRONG("对不起，已经有十个桌面便签被打开，为保证程序运行稳定，您不可以继续添加便签啦", 2): Exit Sub
Dim NBQ As New FrmNew
NBQ.Show
End Sub
Sub LOADCLOUD()
Winsock1.SendData ".getlist"
Debug.Print "请稍等,正在从服务器拉取数据..."
End Sub
Sub LoadInfo()
On Error Resume Next
Open App.Path & "\USER\" & Text1.Text & ".dat" For Random As gFileNum Len = Len(MyPersonalInfo)
Get #gFileNum, 1, MyPersonalInfo
Dim temp As String
MyPersonalInfo.Country = Replace(MyPersonalInfo.Country, " ", "")
MyPersonalInfo.BIRTHDAY = Replace((MyPersonalInfo.BIRTHDAY), " ", "")
MyPersonalInfo.Age = Replace(MyPersonalInfo.Age, " ", "")
MyPersonalInfo.Webpage = Replace(MyPersonalInfo.Webpage, " ", "")
temp = Replace(Trim(MyPersonalInfo.About), "//crlf\\", vbCrLf)
MyPersonalInfo.About = Replace(temp, " ", "")
LBSG.Caption = MyPersonalInfo.Country
MyPersonalInfo.PHONE = Trim(Replace(MyPersonalInfo.PHONE, " ", ""))
MyPersonalInfo.Address = Trim(Replace(MyPersonalInfo.Address, " ", ""))
MyPersonalInfo.QQ = Trim(Replace(MyPersonalInfo.QQ, " ", ""))
MyPersonalInfo.language = Trim(Replace(MyPersonalInfo.language, " ", ""))
MyPersonalInfo.JOB = Trim(Replace(MyPersonalInfo.JOB, " ", ""))
MyPersonalInfo.STUDY = Trim(Replace(MyPersonalInfo.STUDY, " ", ""))
MyPersonalInfo.SX = Trim(Replace(MyPersonalInfo.SX, " ", ""))
MyPersonalInfo.TEL = Trim(Replace(MyPersonalInfo.TEL, " ", ""))
MyPersonalInfo.OAB = Trim(Replace(MyPersonalInfo.OAB, " ", ""))
MyPersonalInfo.COU = Trim(Replace(MyPersonalInfo.COU, " ", ""))
MyPersonalInfo.BIRTH = Trim(Replace(MyPersonalInfo.BIRTH, " ", ""))
Close #gFileNum
End Sub
