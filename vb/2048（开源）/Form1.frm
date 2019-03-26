VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "2048"
   ClientHeight    =   7170
   ClientLeft      =   7605
   ClientTop       =   1605
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7170
   ScaleWidth      =   6360
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6720
      TabIndex        =   25
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10920
      TabIndex        =   23
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "↑"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7800
      TabIndex        =   22
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "←"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   8280
      TabIndex        =   21
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "↓"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9240
      TabIndex        =   20
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command101 
      Caption         =   "分数"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   18
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "→"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   8760
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command66 
      Caption         =   "Command66"
      Enabled         =   0   'False
      Height          =   540
      Left            =   10800
      TabIndex        =   0
      Top             =   7920
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1350
      Index           =   16
      Left            =   4680
      Picture         =   "Form1.frx":1CCD
      Top             =   5520
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   1350
      Index           =   15
      Left            =   3240
      Picture         =   "Form1.frx":1D93
      Top             =   5520
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   1350
      Index           =   14
      Left            =   1800
      Picture         =   "Form1.frx":1E59
      Top             =   5520
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   1350
      Index           =   13
      Left            =   360
      Picture         =   "Form1.frx":1F1F
      Top             =   5520
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   1350
      Index           =   12
      Left            =   4680
      Picture         =   "Form1.frx":1FE5
      Top             =   4080
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   1350
      Index           =   11
      Left            =   3240
      Picture         =   "Form1.frx":20AB
      Top             =   4080
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   1350
      Index           =   10
      Left            =   1800
      Picture         =   "Form1.frx":2171
      Top             =   4080
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   1350
      Index           =   9
      Left            =   360
      Picture         =   "Form1.frx":2237
      Top             =   4080
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   1350
      Index           =   8
      Left            =   4680
      Picture         =   "Form1.frx":22FD
      Top             =   2640
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   1350
      Index           =   7
      Left            =   3240
      Picture         =   "Form1.frx":23C3
      Top             =   2640
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   1350
      Index           =   6
      Left            =   1800
      Picture         =   "Form1.frx":2489
      Top             =   2640
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   1350
      Index           =   5
      Left            =   360
      Picture         =   "Form1.frx":254F
      Top             =   2640
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   1350
      Index           =   4
      Left            =   4680
      Picture         =   "Form1.frx":2615
      Top             =   1200
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   1350
      Index           =   3
      Left            =   3240
      Picture         =   "Form1.frx":26DB
      Top             =   1200
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   1350
      Index           =   2
      Left            =   1800
      Picture         =   "Form1.frx":27A1
      Top             =   1200
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   1350
      Index           =   1
      Left            =   360
      Picture         =   "Form1.frx":2867
      Top             =   1200
      Width           =   1350
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10920
      TabIndex        =   24
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label101 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   19
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   16
      Left            =   9600
      TabIndex        =   17
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   15
      Left            =   8880
      TabIndex        =   16
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   14
      Left            =   8160
      TabIndex        =   15
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   13
      Left            =   7440
      TabIndex        =   14
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   9600
      TabIndex        =   13
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   8880
      TabIndex        =   12
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   8160
      TabIndex        =   11
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   7440
      TabIndex        =   10
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   9600
      TabIndex        =   9
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   8880
      TabIndex        =   8
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   8160
      TabIndex        =   7
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   7440
      TabIndex        =   6
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   9600
      TabIndex        =   5
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   8880
      TabIndex        =   4
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   7440
      TabIndex        =   3
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   8160
      TabIndex        =   2
      Top             =   4680
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a01, a02, a03, a04, a05, a06, a07, a08, a09, a10, a11, a12, a13, a14, a15, a16, S, X, Y, Z, i As Long
Private Sub Command1_Click()
Command66_Click

If a01 <> 0 And a02 = 0 And a03 = 0 And a04 = 0 Then
    Label1(4).Caption = a01
    Label1(1).Caption = 0
End If

If a01 = 0 And a02 <> 0 And a03 = 0 And a04 = 0 Then
    Label1(4).Caption = a02
    Label1(2).Caption = 0
End If

If a01 = 0 And a02 = 0 And a03 <> 0 And a04 = 0 Then
    Label1(4).Caption = a03
    Label1(3).Caption = 0
End If

If a01 = 0 And a02 = 0 And a03 = 0 And a04 <> 0 Then

End If

'————华丽的分割线————

If a01 <> 0 And a02 <> 0 And a03 = 0 And a04 = 0 Then
    If a01 = a02 Then
        Label1(4).Caption = 2 * a01
        S = S + 2 * a01
        Label1(1).Caption = 0
        Label1(2).Caption = 0
    Else
        Label1(3).Caption = a01
        Label1(4).Caption = a02
        Label1(1).Caption = 0
        Label1(2).Caption = 0
    End If
End If

If a01 = 0 And a02 <> 0 And a03 <> 0 And a04 = 0 Then
    If a02 = a03 Then
        Label1(4).Caption = 2 * a02
        S = S + 2 * a02
        Label1(2).Caption = 0
        Label1(3).Caption = 0
    Else
        Label1(4).Caption = a03
        Label1(3).Caption = a02
        Label1(2).Caption = 0
    End If
End If


If a01 = 0 And a02 = 0 And a03 <> 0 And a04 <> 0 Then
    If a03 = a04 Then
        Label1(4).Caption = 2 * a03
        S = S + 2 * a03
        Label1(3).Caption = 0
    End If
End If

If a01 <> 0 And a02 = 0 And a03 <> 0 And a04 = 0 Then
    If a01 = a03 Then
        Label1(4).Caption = 2 * a01
        S = S + 2 * a01
        Label1(1).Caption = 0
        Label1(3).Caption = 0
    Else
        Label1(4).Caption = a03
        Label1(3).Caption = a01
        Label1(1).Caption = 0
    End If
End If

If a01 = 0 And a02 <> 0 And a03 = 0 And a04 <> 0 Then
    If a02 = a04 Then
        Label1(4).Caption = 2 * a02
        S = S + 2 * a02
        Label1(2).Caption = 0
    Else
        Label1(3).Caption = a02
        Label1(2).Caption = 0
    End If
End If

If a01 <> 0 And a02 = 0 And a03 = 0 And a04 <> 0 Then
    If a01 = a04 Then
        Label1(4).Caption = 2 * a01
        S = S + 2 * a01
        Label1(1).Caption = 0
    Else
        Label1(3).Caption = a01
        Label1(1).Caption = 0
    End If
End If
'————华丽的分割线————
If a01 <> 0 And a02 <> 0 And a03 <> 0 And a04 = 0 Then
    If a01 = a02 And a02 <> a03 Then
        Label1(4).Caption = a03
        Label1(3).Caption = 2 * a01
        S = S + 2 * a01
        Label1(2).Caption = 0
        Label1(1).Caption = 0
    Else
        If a02 = a03 And a01 <> a02 Then
            Label1(4).Caption = 2 * a03
            S = S + 2 * a03
            Label1(3).Caption = a01
            Label1(2).Caption = 0
            Label1(1).Caption = 0
        Else
            If a01 = a02 And a01 = a03 Then
                Label1(4).Caption = 2 * a03
                S = S + 2 * a03
                Label1(3).Caption = a01
                Label1(2).Caption = 0
                Label1(1).Caption = 0
            Else
                Label1(4).Caption = a03
                Label1(3).Caption = a02
                Label1(2).Caption = a01
                Label1(1).Caption = 0
            End If
        End If
    End If
End If
'————华丽的分割线————
If a01 = 0 And a02 <> 0 And a03 <> 0 And a04 <> 0 Then
    If a02 = a03 And a03 <> a04 Then
        Label1(3).Caption = 2 * a02
        S = S + 2 * a02
        Label1(2).Caption = 0
        Label1(1).Caption = 0
    Else
        If a03 = a04 And a02 <> a03 Then
            Label1(4).Caption = 2 * a03
            S = S + 2 * a03
            Label1(3).Caption = a02
            Label1(2).Caption = 0
            Label1(1).Caption = 0
        Else
            If a02 = a03 And a02 = a04 Then
                Label1(4).Caption = 2 * a03
                S = S + 2 * a03
                Label1(3).Caption = a02
                Label1(2).Caption = 0
                Label1(1).Caption = 0
            End If
        End If
    End If
End If
'————华丽的分割线————
If a01 <> 0 And a02 <> 0 And a03 = 0 And a04 <> 0 Then
    If a01 = a02 And a01 <> a04 Then
        Label1(3).Caption = 2 * a02
        S = S + 2 * a02
        Label1(2).Caption = 0
        Label1(1).Caption = 0
    Else
        If a01 <> a02 And a02 = a04 Then
            Label1(4).Caption = 2 * a02
            S = S + 2 * a02
            Label1(3).Caption = a01
            Label1(2).Caption = 0
            Label1(1).Caption = 0
        Else
            If a01 = a02 And a02 = a04 Then
                Label1(4).Caption = 2 * a02
                S = S + 2 * a02
                Label1(3).Caption = a01
                Label1(2).Caption = 0
                Label1(1).Caption = 0
            Else
                Label1(3).Caption = a02
                Label1(2).Caption = a01
                Label1(1).Caption = 0
            End If
        End If
    End If
End If
'————华丽的分割线————
If a01 <> 0 And a02 = 0 And a03 <> 0 And a04 <> 0 Then
    If a01 = a03 And a01 <> a04 Then
        Label1(3).Caption = 2 * a01
        S = S + 2 * a01
        Label1(2).Caption = 0
        Label1(1).Caption = 0
    Else
        If a01 <> a03 And a03 = a04 Then
            Label1(4).Caption = 2 * a03
            S = S + 2 * a03
            Label1(3).Caption = a01
            Label1(2).Caption = 0
            Label1(1).Caption = 0
        Else
            If a01 = a03 And a03 = a04 Then
                Label1(4).Caption = 2 * a03
                S = S + 2 * a03
                Label1(3).Caption = a01
                Label1(2).Caption = 0
                Label1(1).Caption = 0
            Else
                Label1(2).Caption = a01
                Label1(1).Caption = 0
            End If
        End If
    End If
End If
'————华丽的分割线————
If a01 <> 0 And a02 <> 0 And a03 <> 0 And a04 <> 0 Then
    If a01 = a02 And a03 = a04 Then
         Label1(4).Caption = 2 * a03
         S = S + 2 * a03
         Label1(3).Caption = 2 * a01
         S = S + 2 * a01
         Label1(2).Caption = 0
         Label1(1).Caption = 0
    Else
        If a01 <> a02 And a02 = a03 And a03 <> a04 Then
            Label1(3).Caption = 2 * a02
            S = S + 2 * a02
            Label1(2).Caption = a01
            Label1(1).Caption = 0
        Else
            If a01 = a02 And a02 <> a03 Then
                Label1(2).Caption = 2 * a01
                S = S + 2 * a01
                Label1(1).Caption = 0
            Else
                If a02 <> a03 And a03 = a04 Then
                    Label1(4).Caption = 2 * a03
                    S = S + 2 * a03
                    Label1(3).Caption = a02
                    Label1(2).Caption = a01
                    Label1(1).Caption = 0
                Else
                    If a01 = a02 And a02 = a03 And a03 <> a04 Then
                        Label1(3).Caption = 2 * a02
                        S = S + 2 * a02
                        Label1(2).Caption = a01
                        Label1(1).Caption = 0
                    Else
                        If a01 <> a02 And a02 = a03 And a03 = a04 Then
                            Label1(4).Caption = 2 * a03
                            S = S + 2 * a03
                            Label1(3).Caption = a02
                            Label1(2).Caption = a01
                            Label1(1).Caption = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
'————————超级华丽的分割线————————
'————————超级华丽的分割线————————
If a05 <> 0 And a06 = 0 And a07 = 0 And a08 = 0 Then
    Label1(8).Caption = a05
    Label1(5).Caption = 0
End If

If a05 = 0 And a06 <> 0 And a07 = 0 And a08 = 0 Then
    Label1(8).Caption = a06
    Label1(6).Caption = 0
End If

If a05 = 0 And a06 = 0 And a07 <> 0 And a08 = 0 Then
    Label1(8).Caption = a07
    Label1(7).Caption = 0
End If

If a05 = 0 And a06 = 0 And a07 = 0 And a08 <> 0 Then

End If

'----华丽的分割线----

If a05 <> 0 And a06 <> 0 And a07 = 0 And a08 = 0 Then
    If a05 = a06 Then
        Label1(8).Caption = 2 * a05
        S = S + 2 * a05
        Label1(5).Caption = 0
        Label1(6).Caption = 0
    Else
        Label1(7).Caption = a05
        Label1(8).Caption = a06
        Label1(5).Caption = 0
        Label1(6).Caption = 0
    End If
End If

If a05 = 0 And a06 <> 0 And a07 <> 0 And a08 = 0 Then
    If a06 = a07 Then
        Label1(8).Caption = 2 * a06
        S = S + 2 * a06
        Label1(6).Caption = 0
        Label1(7).Caption = 0
    Else
        Label1(8).Caption = a07
        Label1(7).Caption = a06
        Label1(6).Caption = 0
    End If
End If

If a05 = 0 And a06 = 0 And a07 <> 0 And a08 <> 0 Then
    If a07 = a08 Then
        Label1(8).Caption = 2 * a07
        S = S + 2 * a07
        Label1(7).Caption = 0
    End If
End If

If a05 <> 0 And a06 = 0 And a07 <> 0 And a08 = 0 Then
    If a05 = a07 Then
        Label1(8).Caption = 2 * a05
        S = S + 2 * a05
        Label1(5).Caption = 0
        Label1(7).Caption = 0
    Else
        Label1(8).Caption = a07
        Label1(7).Caption = a05
        Label1(5).Caption = 0
    End If
End If

If a05 = 0 And a06 <> 0 And a07 = 0 And a08 <> 0 Then
    If a06 = a08 Then
        Label1(8).Caption = 2 * a06
        S = S + 2 * a06
        Label1(6).Caption = 0
    Else
        Label1(7).Caption = a06
        Label1(6).Caption = 0
    End If
End If

If a05 <> 0 And a06 = 0 And a07 = 0 And a08 <> 0 Then
    If a05 = a08 Then
        Label1(8).Caption = 2 * a05
        S = S + 2 * a05
        Label1(5).Caption = 0
    Else
        Label1(7).Caption = a05
        Label1(5).Caption = 0
    End If
End If
'----华丽的分割线----
If a05 <> 0 And a06 <> 0 And a07 <> 0 And a08 = 0 Then
    If a05 = a06 And a06 <> a07 Then
        Label1(8).Caption = a07
        Label1(7).Caption = 2 * a05
        S = S + 2 * a05
        Label1(6).Caption = 0
        Label1(5).Caption = 0
    Else
        If a06 = a07 And a05 <> a06 Then
            Label1(8).Caption = 2 * a07
            S = S + 2 * a07
            Label1(7).Caption = a05
            Label1(6).Caption = 0
            Label1(5).Caption = 0
        Else
            If a05 = a06 And a05 = a07 Then
                Label1(8).Caption = 2 * a07
                S = S + 2 * a07
                Label1(7).Caption = a05
                Label1(6).Caption = 0
                Label1(5).Caption = 0
            Else
                Label1(8).Caption = a07
                Label1(7).Caption = a06
                Label1(6).Caption = a05
                Label1(5).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a05 = 0 And a06 <> 0 And a07 <> 0 And a08 <> 0 Then
    If a06 = a07 And a07 <> a08 Then
        Label1(7).Caption = 2 * a06
        S = S + 2 * a06
        Label1(6).Caption = 0
        Label1(5).Caption = 0
    Else
        If a07 = a08 And a06 <> a07 Then
            Label1(8).Caption = 2 * a07
            S = S + 2 * a07
            Label1(7).Caption = a06
            Label1(6).Caption = 0
            Label1(5).Caption = 0
        Else
            If a06 = a07 And a06 = a08 Then
                Label1(8).Caption = 2 * a07
                S = S + 2 * a07
                Label1(7).Caption = a06
                Label1(6).Caption = 0
                Label1(5).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a05 <> 0 And a06 <> 0 And a07 = 0 And a08 <> 0 Then
    If a05 = a06 And a05 <> a08 Then
        Label1(7).Caption = 2 * a06
        S = S + 2 * a06
        Label1(6).Caption = 0
        Label1(5).Caption = 0
    Else
        If a05 <> a06 And a06 = a08 Then
            Label1(8).Caption = 2 * a06
            S = S + 2 * a06
            Label1(7).Caption = a05
            Label1(6).Caption = 0
            Label1(5).Caption = 0
        Else
            If a05 = a06 And a06 = a08 Then
                Label1(8).Caption = 2 * a06
                S = S + 2 * a06
                Label1(7).Caption = a05
                Label1(6).Caption = 0
                Label1(5).Caption = 0
            Else
                Label1(7).Caption = a06
                Label1(6).Caption = a05
                Label1(5).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a05 <> 0 And a06 = 0 And a07 <> 0 And a08 <> 0 Then
    If a05 = a07 And a05 <> a08 Then
        Label1(7).Caption = 2 * a05
        S = S + 2 * a05
        Label1(6).Caption = 0
        Label1(5).Caption = 0
    Else
        If a05 <> a07 And a07 = a08 Then
            Label1(8).Caption = 2 * a07
            S = S + 2 * a07
            Label1(7).Caption = a05
            Label1(6).Caption = 0
            Label1(5).Caption = 0
        Else
            If a05 = a07 And a07 = a08 Then
                Label1(8).Caption = 2 * a07
                S = S + 2 * a07
                Label1(7).Caption = a05
                Label1(6).Caption = 0
                Label1(5).Caption = 0
            Else
                Label1(6).Caption = a05
                Label1(5).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a05 <> 0 And a06 <> 0 And a07 <> 0 And a08 <> 0 Then
    If a05 = a06 And a07 = a08 Then
         Label1(8).Caption = 2 * a07
         S = S + 2 * a07
         Label1(7).Caption = 2 * a05
         S = S + 2 * a05
         Label1(6).Caption = 0
         Label1(5).Caption = 0
    Else
        If a05 <> a06 And a06 = a07 And a07 <> a08 Then
            Label1(7).Caption = 2 * a06
            S = S + 2 * a06
            Label1(6).Caption = a05
            Label1(5).Caption = 0
        Else
            If a05 = a06 And a06 <> a07 Then
                Label1(6).Caption = 2 * a05
                S = S + 2 * a05
                Label1(5).Caption = 0
            Else
                If a06 <> a07 And a07 = a08 Then
                    Label1(8).Caption = 2 * a07
                    S = S + 2 * a07
                    Label1(7).Caption = a06
                    Label1(6).Caption = a05
                    Label1(5).Caption = 0
                Else
                    If a05 = a06 And a06 = a07 And a07 <> a08 Then
                        Label1(7).Caption = 2 * a06
                        S = S + 2 * a06
                        Label1(6).Caption = a05
                        Label1(5).Caption = 0
                    Else
                        If a05 <> a06 And a06 = a07 And a07 = a08 Then
                            Label1(8).Caption = 2 * a07
                            S = S + 2 * a07
                            Label1(7).Caption = a06
                            Label1(6).Caption = a05
                            Label1(5).Caption = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
'--------超级华丽的分割线--------
'--------超级华丽的分割线--------
Call Command66_Click

If a09 <> 0 And a10 = 0 And a11 = 0 And a12 = 0 Then
    Label1(12).Caption = a09
    Label1(9).Caption = 0
End If

If a09 = 0 And a10 <> 0 And a11 = 0 And a12 = 0 Then
    Label1(12).Caption = a10
    Label1(10).Caption = 0
End If

If a09 = 0 And a10 = 0 And a11 <> 0 And a12 = 0 Then
    Label1(12).Caption = a11
    Label1(11).Caption = 0
End If

If a09 = 0 And a10 = 0 And a11 = 0 And a12 <> 0 Then

End If

'----华丽的分割线----

If a09 <> 0 And a10 <> 0 And a11 = 0 And a12 = 0 Then
    If a09 = a10 Then
        Label1(12).Caption = 2 * a09
        S = S + 2 * a09
        Label1(9).Caption = 0
        Label1(10).Caption = 0
    Else
        Label1(11).Caption = a09
        Label1(12).Caption = a10
        Label1(9).Caption = 0
        Label1(10).Caption = 0
    End If
End If

If a09 = 0 And a10 <> 0 And a11 <> 0 And a12 = 0 Then
    If a10 = a11 Then
        Label1(12).Caption = 2 * a10
        S = S + 2 * a10
        Label1(10).Caption = 0
        Label1(11).Caption = 0
    Else
        Label1(12).Caption = a11
        Label1(11).Caption = a10
        Label1(10).Caption = 0
    End If
End If

If a09 = 0 And a10 = 0 And a11 <> 0 And a12 <> 0 Then
    If a11 = a12 Then
        Label1(12).Caption = 2 * a11
        S = S + 2 * a11
        Label1(11).Caption = 0
    End If
End If

If a09 <> 0 And a10 = 0 And a11 <> 0 And a12 = 0 Then
    If a09 = a11 Then
        Label1(12).Caption = 2 * a09
        S = S + 2 * a09
        Label1(9).Caption = 0
        Label1(11).Caption = 0
    Else
        Label1(12).Caption = a11
        Label1(11).Caption = a09
        Label1(9).Caption = 0
    End If
End If

If a09 = 0 And a10 <> 0 And a11 = 0 And a12 <> 0 Then
    If a10 = a12 Then
        Label1(12).Caption = 2 * a10
        S = S + 2 * a10
        Label1(10).Caption = 0
    Else
        Label1(11).Caption = a10
        Label1(10).Caption = 0
    End If
End If

If a09 <> 0 And a10 = 0 And a11 = 0 And a12 <> 0 Then
    If a09 = a12 Then
        Label1(12).Caption = 2 * a09
        S = S + 2 * a09
        Label1(9).Caption = 0
    Else
        Label1(11).Caption = a09
        Label1(9).Caption = 0
    End If
End If
'----华丽的分割线----
If a09 <> 0 And a10 <> 0 And a11 <> 0 And a12 = 0 Then
    If a09 = a10 And a10 <> a11 Then
        Label1(12).Caption = a11
        Label1(11).Caption = 2 * a09
        S = S + 2 * a09
        Label1(10).Caption = 0
        Label1(9).Caption = 0
    Else
        If a10 = a11 And a09 <> a10 Then
            Label1(12).Caption = 2 * a11
            S = S + 2 * a11
            Label1(11).Caption = a09
            Label1(10).Caption = 0
            Label1(9).Caption = 0
        Else
            If a09 = a10 And a09 = a11 Then
                Label1(12).Caption = 2 * a11
                S = S + 2 * a11
                Label1(11).Caption = a09
                Label1(10).Caption = 0
                Label1(9).Caption = 0
            Else
                Label1(12).Caption = a11
                Label1(11).Caption = a10
                Label1(10).Caption = a09
                Label1(9).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a09 = 0 And a10 <> 0 And a11 <> 0 And a12 <> 0 Then
    If a10 = a11 And a11 <> a12 Then
        Label1(11).Caption = 2 * a10
        S = S + 2 * a10
        Label1(10).Caption = 0
        Label1(9).Caption = 0
    Else
        If a11 = a12 And a10 <> a11 Then
            Label1(12).Caption = 2 * a11
            S = S + 2 * a11
            Label1(11).Caption = a10
            Label1(10).Caption = 0
            Label1(9).Caption = 0
        Else
            If a10 = a11 And a10 = a12 Then
                Label1(12).Caption = 2 * a11
                S = S + 2 * a11
                Label1(11).Caption = a10
                Label1(10).Caption = 0
                Label1(9).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a09 <> 0 And a10 <> 0 And a11 = 0 And a12 <> 0 Then
    If a09 = a10 And a09 <> a12 Then
        Label1(11).Caption = 2 * a10
        S = S + 2 * a10
        Label1(10).Caption = 0
        Label1(9).Caption = 0
    Else
        If a09 <> a10 And a10 = a12 Then
            Label1(12).Caption = 2 * a10
            S = S + 2 * a10
            Label1(11).Caption = a09
            Label1(10).Caption = 0
            Label1(9).Caption = 0
        Else
            If a09 = a10 And a10 = a12 Then
                Label1(12).Caption = 2 * a10
                S = S + 2 * a10
                Label1(11).Caption = a09
                Label1(10).Caption = 0
                Label1(9).Caption = 0
            Else
                Label1(11).Caption = a10
                Label1(10).Caption = a09
                Label1(9).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a09 <> 0 And a10 = 0 And a11 <> 0 And a12 <> 0 Then
    If a09 = a11 And a09 <> a12 Then
        Label1(11).Caption = 2 * a09
        S = S + 2 * a09
        Label1(10).Caption = 0
        Label1(9).Caption = 0
    Else
        If a09 <> a11 And a11 = a12 Then
            Label1(12).Caption = 2 * a11
            S = S + 2 * a11
            Label1(11).Caption = a09
            Label1(10).Caption = 0
            Label1(9).Caption = 0
        Else
            If a09 = a11 And a11 = a12 Then
                Label1(12).Caption = 2 * a11
                S = S + 2 * a11
                Label1(11).Caption = a09
                Label1(10).Caption = 0
                Label1(9).Caption = 0
            Else
                Label1(10).Caption = a09
                Label1(9).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a09 <> 0 And a10 <> 0 And a11 <> 0 And a12 <> 0 Then
    If a09 = a10 And a11 = a12 Then
         Label1(12).Caption = 2 * a11
         S = S + 2 * a11
         Label1(11).Caption = 2 * a09
         S = S + 2 * a09
         Label1(10).Caption = 0
         Label1(9).Caption = 0
    Else
        If a09 <> a10 And a10 = a11 And a11 <> a12 Then
            Label1(11).Caption = 2 * a10
            S = S + 2 * a10
            Label1(10).Caption = a09
            Label1(9).Caption = 0
        Else
            If a09 = a10 And a10 <> a11 Then
                Label1(10).Caption = 2 * a09
                S = S + 2 * a09
                Label1(9).Caption = 0
            Else
                If a10 <> a11 And a11 = a12 Then
                    Label1(12).Caption = 2 * a11
                    S = S + 2 * a11
                    Label1(11).Caption = a10
                    Label1(10).Caption = a09
                    Label1(9).Caption = 0
                Else
                    If a09 = a10 And a10 = a11 And a11 <> a12 Then
                        Label1(11).Caption = 2 * a10
                        S = S + 2 * a10
                        Label1(10).Caption = a09
                        Label1(9).Caption = 0
                    Else
                        If a09 <> a10 And a10 = a11 And a11 = a12 Then
                            Label1(12).Caption = 2 * a11
                            S = S + 2 * a11
                            Label1(11).Caption = a10
                            Label1(10).Caption = a09
                            Label1(9).Caption = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
'--------超级华丽的分割线--------
'--------超级华丽的分割线--------
If a13 <> 0 And a14 = 0 And a15 = 0 And a16 = 0 Then
    Label1(16).Caption = a13
    Label1(13).Caption = 0
End If

If a13 = 0 And a14 <> 0 And a15 = 0 And a16 = 0 Then
    Label1(16).Caption = a14
    Label1(14).Caption = 0
End If

If a13 = 0 And a14 = 0 And a15 <> 0 And a16 = 0 Then
    Label1(16).Caption = a15
    Label1(15).Caption = 0
End If

If a13 = 0 And a14 = 0 And a15 = 0 And a16 <> 0 Then

End If

'----华丽的分割线----

If a13 <> 0 And a14 <> 0 And a15 = 0 And a16 = 0 Then
    If a13 = a14 Then
        Label1(16).Caption = 2 * a13
        S = S + 2 * a13
        Label1(13).Caption = 0
        Label1(14).Caption = 0
    Else
        Label1(15).Caption = a13
        Label1(16).Caption = a14
        Label1(13).Caption = 0
        Label1(14).Caption = 0
    End If
End If

If a13 = 0 And a14 <> 0 And a15 <> 0 And a16 = 0 Then
    If a14 = a15 Then
        Label1(16).Caption = 2 * a14
        S = S + 2 * a14
        Label1(14).Caption = 0
        Label1(15).Caption = 0
    Else
        Label1(16).Caption = a15
        Label1(15).Caption = a14
        Label1(14).Caption = 0
    End If
End If

If a13 = 0 And a14 = 0 And a15 <> 0 And a16 <> 0 Then
    If a15 = a16 Then
        Label1(16).Caption = 2 * a15
        S = S + 2 * a15
        Label1(15).Caption = 0
    End If
End If

If a13 <> 0 And a14 = 0 And a15 <> 0 And a16 = 0 Then
    If a13 = a15 Then
        Label1(16).Caption = 2 * a13
        S = S + 2 * a13
        Label1(13).Caption = 0
        Label1(15).Caption = 0
    Else
        Label1(16).Caption = a15
        Label1(15).Caption = a13
        Label1(13).Caption = 0
    End If
End If

If a13 = 0 And a14 <> 0 And a15 = 0 And a16 <> 0 Then
    If a14 = a16 Then
        Label1(16).Caption = 2 * a14
        S = S + 2 * a14
        Label1(14).Caption = 0
    Else
        Label1(15).Caption = a14
        Label1(14).Caption = 0
    End If
End If

If a13 <> 0 And a14 = 0 And a15 = 0 And a16 <> 0 Then
    If a13 = a16 Then
        Label1(16).Caption = 2 * a13
        S = S + 2 * a13
        Label1(13).Caption = 0
    Else
        Label1(15).Caption = a13
        Label1(13).Caption = 0
    End If
End If
'----华丽的分割线----
If a13 <> 0 And a14 <> 0 And a15 <> 0 And a16 = 0 Then
    If a13 = a14 And a14 <> a15 Then
        Label1(16).Caption = a15
        Label1(15).Caption = 2 * a13
        S = S + 2 * a13
        Label1(14).Caption = 0
        Label1(13).Caption = 0
    Else
        If a14 = a15 And a13 <> a14 Then
            Label1(16).Caption = 2 * a15
            S = S + 2 * a15
            Label1(15).Caption = a13
            Label1(14).Caption = 0
            Label1(13).Caption = 0
        Else
            If a13 = a14 And a13 = a15 Then
                Label1(16).Caption = 2 * a15
                S = S + 2 * a15
                Label1(15).Caption = a13
                Label1(14).Caption = 0
                Label1(13).Caption = 0
            Else
                Label1(16).Caption = a15
                Label1(15).Caption = a14
                Label1(14).Caption = a13
                Label1(13).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a13 = 0 And a14 <> 0 And a15 <> 0 And a16 <> 0 Then
    If a14 = a15 And a15 <> a16 Then
        Label1(15).Caption = 2 * a14
        S = S + 2 * a14
        Label1(14).Caption = 0
        Label1(13).Caption = 0
    Else
        If a15 = a16 And a14 <> a15 Then
            Label1(16).Caption = 2 * a15
            S = S + 2 * a15
            Label1(15).Caption = a14
            Label1(14).Caption = 0
            Label1(13).Caption = 0
        Else
            If a14 = a15 And a14 = a16 Then
                Label1(16).Caption = 2 * a15
                S = S + 2 * a15
                Label1(15).Caption = a14
                Label1(14).Caption = 0
                Label1(13).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a13 <> 0 And a14 <> 0 And a15 = 0 And a16 <> 0 Then
    If a13 = a14 And a13 <> a16 Then
        Label1(15).Caption = 2 * a14
        S = S + 2 * a14
        Label1(14).Caption = 0
        Label1(13).Caption = 0
    Else
        If a13 <> a14 And a14 = a16 Then
            Label1(16).Caption = 2 * a14
            S = S + 2 * a14
            Label1(15).Caption = a13
            Label1(14).Caption = 0
            Label1(13).Caption = 0
        Else
            If a13 = a14 And a14 = a16 Then
                Label1(16).Caption = 2 * a14
                S = S + 2 * a14
                Label1(15).Caption = a13
                Label1(14).Caption = 0
                Label1(13).Caption = 0
            Else
                Label1(15).Caption = a14
                Label1(14).Caption = a13
                Label1(13).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a13 <> 0 And a14 = 0 And a15 <> 0 And a16 <> 0 Then
    If a13 = a15 And a13 <> a16 Then
        Label1(15).Caption = 2 * a13
        S = S + 2 * a13
        Label1(14).Caption = 0
        Label1(13).Caption = 0
    Else
        If a13 <> a15 And a15 = a16 Then
            Label1(16).Caption = 2 * a15
            S = S + 2 * a15
            Label1(15).Caption = a13
            Label1(14).Caption = 0
            Label1(13).Caption = 0
        Else
            If a13 = a15 And a15 = a16 Then
                Label1(16).Caption = 2 * a15
                S = S + 2 * a15
                Label1(15).Caption = a13
                Label1(14).Caption = 0
                Label1(13).Caption = 0
            Else
                Label1(14).Caption = a13
                Label1(13).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a13 <> 0 And a14 <> 0 And a15 <> 0 And a16 <> 0 Then
    If a13 = a14 And a15 = a16 Then
         Label1(16).Caption = 2 * a15
         S = S + 2 * a15
         Label1(15).Caption = 2 * a13
         S = S + 2 * a13
         Label1(14).Caption = 0
         Label1(13).Caption = 0
    Else
        If a13 <> a14 And a14 = a15 And a15 <> a16 Then
            Label1(15).Caption = 2 * a14
            S = S + 2 * a14
            Label1(14).Caption = a13
            Label1(13).Caption = 0
        Else
            If a13 = a14 And a14 <> a15 Then
                Label1(14).Caption = 2 * a13
                S = S + 2 * a13
                Label1(13).Caption = 0
            Else
                If a14 <> a15 And a15 = a16 Then
                    Label1(16).Caption = 2 * a15
                    S = S + 2 * a15
                    Label1(15).Caption = a14
                    Label1(14).Caption = a13
                    Label1(13).Caption = 0
                Else
                    If a13 = a14 And a14 = a15 And a15 <> a16 Then
                        Label1(15).Caption = 2 * a14
                        S = S + 2 * a14
                        Label1(14).Caption = a13
                        Label1(13).Caption = 0
                    Else
                        If a13 <> a14 And a14 = a15 And a15 = a16 Then
                            Label1(16).Caption = 2 * a15
                            S = S + 2 * a15
                            Label1(15).Caption = a14
                            Label1(14).Caption = a13
                            Label1(13).Caption = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
Call Command101_Click
Call Command5_Click
Call Command6_Click
End Sub



Private Sub Command101_Click()
Label101.Caption = S
End Sub

Private Sub Command2_Click()
Call Command66_Click

If a01 <> 0 And a05 = 0 And a09 = 0 And a13 = 0 Then
    Label1(13).Caption = a01
    Label1(1).Caption = 0
End If

If a01 = 0 And a05 <> 0 And a09 = 0 And a13 = 0 Then
    Label1(13).Caption = a05
    Label1(5).Caption = 0
End If

If a01 = 0 And a05 = 0 And a09 <> 0 And a13 = 0 Then
    Label1(13).Caption = a09
    Label1(9).Caption = 0
End If

If a01 = 0 And a05 = 0 And a09 = 0 And a13 <> 0 Then

End If

'----华丽的分割线----

If a01 <> 0 And a05 <> 0 And a09 = 0 And a13 = 0 Then
    If a01 = a05 Then
        Label1(13).Caption = 2 * a01
        S = S + 2 * a01
        Label1(1).Caption = 0
        Label1(5).Caption = 0
    Else
        Label1(9).Caption = a01
        Label1(13).Caption = a05
        Label1(1).Caption = 0
        Label1(5).Caption = 0
    End If
End If

If a01 = 0 And a05 <> 0 And a09 <> 0 And a13 = 0 Then
    If a05 = a09 Then
        Label1(13).Caption = 2 * a05
        S = S + 2 * a05
        Label1(5).Caption = 0
        Label1(9).Caption = 0
    Else
        Label1(13).Caption = a09
        Label1(9).Caption = a05
        Label1(5).Caption = 0
    End If
End If

If a01 = 0 And a05 = 0 And a09 <> 0 And a13 <> 0 Then
    If a09 = a13 Then
        Label1(13).Caption = 2 * a09
        S = S + 2 * a09
        Label1(9).Caption = 0
    End If
End If

If a01 <> 0 And a05 = 0 And a09 <> 0 And a13 = 0 Then
    If a01 = a09 Then
        Label1(13).Caption = 2 * a01
        S = S + 2 * a01
        Label1(1).Caption = 0
        Label1(9).Caption = 0
    Else
        Label1(13).Caption = a09
        Label1(9).Caption = a01
        Label1(1).Caption = 0
    End If
End If

If a01 = 0 And a05 <> 0 And a09 = 0 And a13 <> 0 Then
    If a05 = a13 Then
        Label1(13).Caption = 2 * a05
        S = S + 2 * a05
        Label1(5).Caption = 0
    Else
        Label1(9).Caption = a05
        Label1(5).Caption = 0
    End If
End If

If a01 <> 0 And a05 = 0 And a09 = 0 And a13 <> 0 Then
    If a01 = a13 Then
        Label1(13).Caption = 2 * a01
        S = S + 2 * a01
        Label1(1).Caption = 0
    Else
        Label1(9).Caption = a01
        Label1(1).Caption = 0
    End If
End If
'----华丽的分割线----
If a01 <> 0 And a05 <> 0 And a09 <> 0 And a13 = 0 Then
    If a01 = a05 And a05 <> a09 Then
        Label1(13).Caption = a09
        Label1(9).Caption = 2 * a01
        S = S + 2 * a01
        Label1(5).Caption = 0
        Label1(1).Caption = 0
    Else
        If a05 = a09 And a01 <> a05 Then
            Label1(13).Caption = 2 * a09
            S = S + 2 * a09
            Label1(9).Caption = a01
            Label1(5).Caption = 0
            Label1(1).Caption = 0
        Else
            If a01 = a05 And a01 = a09 Then
                Label1(13).Caption = 2 * a09
                S = S + 2 * a09
                Label1(9).Caption = a01
                Label1(5).Caption = 0
                Label1(1).Caption = 0
            Else
                Label1(13).Caption = a09
                Label1(9).Caption = a05
                Label1(5).Caption = a01
                Label1(1).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a01 = 0 And a05 <> 0 And a09 <> 0 And a13 <> 0 Then
    If a05 = a09 And a09 <> a13 Then
        Label1(9).Caption = 2 * a05
        S = S + 2 * a05
        Label1(5).Caption = 0
        Label1(1).Caption = 0
    Else
        If a09 = a13 And a05 <> a09 Then
            Label1(13).Caption = 2 * a09
            S = S + 2 * a09
            Label1(9).Caption = a05
            Label1(5).Caption = 0
            Label1(1).Caption = 0
        Else
            If a05 = a09 And a05 = a13 Then
                Label1(13).Caption = 2 * a09
                S = S + 2 * a09
                Label1(9).Caption = a05
                Label1(5).Caption = 0
                Label1(1).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a01 <> 0 And a05 <> 0 And a09 = 0 And a13 <> 0 Then
    If a01 = a05 And a01 <> a13 Then
        Label1(9).Caption = 2 * a05
        S = S + 2 * a05
        Label1(5).Caption = 0
        Label1(1).Caption = 0
    Else
        If a01 <> a05 And a05 = a13 Then
            Label1(13).Caption = 2 * a05
            S = S + 2 * a05
            Label1(9).Caption = a01
            Label1(5).Caption = 0
            Label1(1).Caption = 0
        Else
            If a01 = a05 And a05 = a13 Then
                Label1(13).Caption = 2 * a05
                S = S + 2 * a05
                Label1(9).Caption = a01
                Label1(5).Caption = 0
                Label1(1).Caption = 0
            Else
                Label1(9).Caption = a05
                Label1(5).Caption = a01
                Label1(1).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a01 <> 0 And a05 = 0 And a09 <> 0 And a13 <> 0 Then
    If a01 = a09 And a01 <> a13 Then
        Label1(9).Caption = 2 * a01
        S = S + 2 * a01
        Label1(5).Caption = 0
        Label1(1).Caption = 0
    Else
        If a01 <> a09 And a09 = a13 Then
            Label1(13).Caption = 2 * a09
            S = S + 2 * a09
            Label1(9).Caption = a01
            Label1(5).Caption = 0
            Label1(1).Caption = 0
        Else
            If a01 = a09 And a09 = a13 Then
                Label1(13).Caption = 2 * a09
                S = S + 2 * a09
                Label1(9).Caption = a01
                Label1(5).Caption = 0
                Label1(1).Caption = 0
            Else
                Label1(5).Caption = a01
                Label1(1).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a01 <> 0 And a05 <> 0 And a09 <> 0 And a13 <> 0 Then
    If a01 = a05 And a09 = a13 Then
         Label1(13).Caption = 2 * a09
         S = S + 2 * a09
         Label1(9).Caption = 2 * a01
         S = S + 2 * a01
         Label1(5).Caption = 0
         Label1(1).Caption = 0
    Else
        If a01 <> a05 And a05 = a09 And a09 <> a13 Then
            Label1(9).Caption = 2 * a05
            S = S + 2 * a05
            Label1(5).Caption = a01
            Label1(1).Caption = 0
        Else
            If a01 = a05 And a05 <> a09 Then
                Label1(5).Caption = 2 * a01
                S = S + 2 * a01
                Label1(1).Caption = 0
            Else
                If a05 <> a09 And a09 = a13 Then
                    Label1(13).Caption = 2 * a09
                    S = S + 2 * a09
                    Label1(9).Caption = a05
                    Label1(5).Caption = a01
                    Label1(1).Caption = 0
                Else
                    If a01 = a05 And a05 = a09 And a09 <> a13 Then
                        Label1(9).Caption = 2 * a05
                        S = S + 2 * a05
                        Label1(5).Caption = a01
                        Label1(1).Caption = 0
                    Else
                        If a01 <> a05 And a05 = a09 And a09 = a13 Then
                            Label1(13).Caption = 2 * a09
                            S = S + 2 * a09
                            Label1(9).Caption = a05
                            Label1(5).Caption = a01
                            Label1(1).Caption = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
'--------超级华丽的分割线--------
'--------超级华丽的分割线--------
If a02 <> 0 And a06 = 0 And a10 = 0 And a14 = 0 Then
    Label1(14).Caption = a02
    Label1(2).Caption = 0
End If

If a02 = 0 And a06 <> 0 And a10 = 0 And a14 = 0 Then
    Label1(14).Caption = a06
    Label1(6).Caption = 0
End If

If a02 = 0 And a06 = 0 And a10 <> 0 And a14 = 0 Then
    Label1(14).Caption = a10
    Label1(10).Caption = 0
End If

If a02 = 0 And a06 = 0 And a10 = 0 And a14 <> 0 Then

End If

'----华丽的分割线----

If a02 <> 0 And a06 <> 0 And a10 = 0 And a14 = 0 Then
    If a02 = a06 Then
        Label1(14).Caption = 2 * a02
        S = S + 2 * a02
        Label1(2).Caption = 0
        Label1(6).Caption = 0
    Else
        Label1(10).Caption = a02
        Label1(14).Caption = a06
        Label1(2).Caption = 0
        Label1(6).Caption = 0
    End If
End If

If a02 = 0 And a06 <> 0 And a10 <> 0 And a14 = 0 Then
    If a06 = a10 Then
        Label1(14).Caption = 2 * a06
        S = S + 2 * a06
        Label1(6).Caption = 0
        Label1(10).Caption = 0
    Else
        Label1(14).Caption = a10
        Label1(10).Caption = a06
        Label1(6).Caption = 0
    End If
End If

If a02 = 0 And a06 = 0 And a10 <> 0 And a14 <> 0 Then
    If a10 = a14 Then
        Label1(14).Caption = 2 * a10
        S = S + 2 * a10
        Label1(10).Caption = 0
    End If
End If

If a02 <> 0 And a06 = 0 And a10 <> 0 And a14 = 0 Then
    If a02 = a10 Then
        Label1(14).Caption = 2 * a02
        S = S + 2 * a02
        Label1(2).Caption = 0
        Label1(10).Caption = 0
    Else
        Label1(14).Caption = a10
        Label1(10).Caption = a02
        Label1(2).Caption = 0
    End If
End If

If a02 = 0 And a06 <> 0 And a10 = 0 And a14 <> 0 Then
    If a06 = a14 Then
        Label1(14).Caption = 2 * a06
        S = S + 2 * a06
        Label1(6).Caption = 0
    Else
        Label1(10).Caption = a06
        Label1(6).Caption = 0
    End If
End If

If a02 <> 0 And a06 = 0 And a10 = 0 And a14 <> 0 Then
    If a02 = a14 Then
        Label1(14).Caption = 2 * a02
        S = S + 2 * a02
        Label1(2).Caption = 0
    Else
        Label1(10).Caption = a02
        Label1(2).Caption = 0
    End If
End If
'----华丽的分割线----
If a02 <> 0 And a06 <> 0 And a10 <> 0 And a14 = 0 Then
    If a02 = a06 And a06 <> a10 Then
        Label1(14).Caption = a10
        Label1(10).Caption = 2 * a02
        S = S + 2 * a02
        Label1(6).Caption = 0
        Label1(2).Caption = 0
    Else
        If a06 = a10 And a02 <> a06 Then
            Label1(14).Caption = 2 * a10
            S = S + 2 * a10
            Label1(10).Caption = a02
            Label1(6).Caption = 0
            Label1(2).Caption = 0
        Else
            If a02 = a06 And a02 = a10 Then
                Label1(14).Caption = 2 * a10
                S = S + 2 * a10
                Label1(10).Caption = a02
                Label1(6).Caption = 0
                Label1(2).Caption = 0
            Else
                Label1(14).Caption = a10
                Label1(10).Caption = a06
                Label1(6).Caption = a02
                Label1(2).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a02 = 0 And a06 <> 0 And a10 <> 0 And a14 <> 0 Then
    If a06 = a10 And a10 <> a14 Then
        Label1(10).Caption = 2 * a06
        S = S + 2 * a06
        Label1(6).Caption = 0
        Label1(2).Caption = 0
    Else
        If a10 = a14 And a06 <> a10 Then
            Label1(14).Caption = 2 * a10
            S = S + 2 * a10
            Label1(10).Caption = a06
            Label1(6).Caption = 0
            Label1(2).Caption = 0
        Else
            If a06 = a10 And a06 = a14 Then
                Label1(14).Caption = 2 * a10
                S = S + 2 * a10
                Label1(10).Caption = a06
                Label1(6).Caption = 0
                Label1(2).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a02 <> 0 And a06 <> 0 And a10 = 0 And a14 <> 0 Then
    If a02 = a06 And a02 <> a14 Then
        Label1(10).Caption = 2 * a06
        S = S + 2 * a06
        Label1(6).Caption = 0
        Label1(2).Caption = 0
    Else
        If a02 <> a06 And a06 = a14 Then
            Label1(14).Caption = 2 * a06
            S = S + 2 * a06
            Label1(10).Caption = a02
            Label1(6).Caption = 0
            Label1(2).Caption = 0
        Else
            If a02 = a06 And a06 = a14 Then
                Label1(14).Caption = 2 * a06
                S = S + 2 * a06
                Label1(10).Caption = a02
                Label1(6).Caption = 0
                Label1(2).Caption = 0
            Else
                Label1(10).Caption = a06
                Label1(6).Caption = a02
                Label1(2).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a02 <> 0 And a06 = 0 And a10 <> 0 And a14 <> 0 Then
    If a02 = a10 And a02 <> a14 Then
        Label1(10).Caption = 2 * a02
        S = S + 2 * a02
        Label1(6).Caption = 0
        Label1(2).Caption = 0
    Else
        If a02 <> a10 And a10 = a14 Then
            Label1(14).Caption = 2 * a10
            S = S + 2 * a10
            Label1(10).Caption = a02
            Label1(6).Caption = 0
            Label1(2).Caption = 0
        Else
            If a02 = a10 And a10 = a14 Then
                Label1(14).Caption = 2 * a10
                S = S + 2 * a10
                Label1(10).Caption = a02
                Label1(6).Caption = 0
                Label1(2).Caption = 0
            Else
                Label1(6).Caption = a02
                Label1(2).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a02 <> 0 And a06 <> 0 And a10 <> 0 And a14 <> 0 Then
    If a02 = a06 And a10 = a14 Then
         Label1(14).Caption = 2 * a10
         S = S + 2 * a10
         Label1(10).Caption = 2 * a02
         S = S + 2 * a02
         Label1(6).Caption = 0
         Label1(2).Caption = 0
    Else
        If a02 <> a06 And a06 = a10 And a10 <> a14 Then
            Label1(10).Caption = 2 * a06
            S = S + 2 * a06
            Label1(6).Caption = a02
            Label1(2).Caption = 0
        Else
            If a02 = a06 And a06 <> a10 Then
                Label1(6).Caption = 2 * a02
                S = S + 2 * a02
                Label1(2).Caption = 0
            Else
                If a06 <> a10 And a10 = a14 Then
                    Label1(14).Caption = 2 * a10
                    S = S + 2 * a10
                    Label1(10).Caption = a06
                    Label1(6).Caption = a02
                    Label1(2).Caption = 0
                Else
                    If a02 = a06 And a06 = a10 And a10 <> a14 Then
                        Label1(10).Caption = 2 * a06
                        S = S + 2 * a06
                        Label1(6).Caption = a02
                        Label1(2).Caption = 0
                    Else
                        If a02 <> a06 And a06 = a10 And a10 = a14 Then
                            Label1(14).Caption = 2 * a10
                            S = S + 2 * a10
                            Label1(10).Caption = a06
                            Label1(6).Caption = a02
                            Label1(2).Caption = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
'--------超级华丽的分割线--------
'--------超级华丽的分割线--------
Call Command66_Click

If a03 <> 0 And a07 = 0 And a11 = 0 And a15 = 0 Then
    Label1(15).Caption = a03
    Label1(3).Caption = 0
End If

If a03 = 0 And a07 <> 0 And a11 = 0 And a15 = 0 Then
    Label1(15).Caption = a07
    Label1(7).Caption = 0
End If

If a03 = 0 And a07 = 0 And a11 <> 0 And a15 = 0 Then
    Label1(15).Caption = a11
    Label1(11).Caption = 0
End If

If a03 = 0 And a07 = 0 And a11 = 0 And a15 <> 0 Then

End If

'----华丽的分割线----

If a03 <> 0 And a07 <> 0 And a11 = 0 And a15 = 0 Then
    If a03 = a07 Then
        Label1(15).Caption = 2 * a03
        S = S + 2 * a03
        Label1(3).Caption = 0
        Label1(7).Caption = 0
    Else
        Label1(11).Caption = a03
        Label1(15).Caption = a07
        Label1(3).Caption = 0
        Label1(7).Caption = 0
    End If
End If

If a03 = 0 And a07 <> 0 And a11 <> 0 And a15 = 0 Then
    If a07 = a11 Then
        Label1(15).Caption = 2 * a07
        S = S + 2 * a07
        Label1(7).Caption = 0
        Label1(11).Caption = 0
    Else
        Label1(15).Caption = a11
        Label1(11).Caption = a07
        Label1(7).Caption = 0
    End If
End If

If a03 = 0 And a07 = 0 And a11 <> 0 And a15 <> 0 Then
    If a11 = a15 Then
        Label1(15).Caption = 2 * a11
        S = S + 2 * a11
        Label1(11).Caption = 0
    End If
End If

If a03 <> 0 And a07 = 0 And a11 <> 0 And a15 = 0 Then
    If a03 = a11 Then
        Label1(15).Caption = 2 * a03
        S = S + 2 * a03
        Label1(3).Caption = 0
        Label1(11).Caption = 0
    Else
        Label1(15).Caption = a11
        Label1(11).Caption = a03
        Label1(3).Caption = 0
    End If
End If

If a03 = 0 And a07 <> 0 And a11 = 0 And a15 <> 0 Then
    If a07 = a15 Then
        Label1(15).Caption = 2 * a07
        S = S + 2 * a07
        Label1(7).Caption = 0
    Else
        Label1(11).Caption = a07
        Label1(7).Caption = 0
    End If
End If

If a03 <> 0 And a07 = 0 And a11 = 0 And a15 <> 0 Then
    If a03 = a15 Then
        Label1(15).Caption = 2 * a03
        S = S + 2 * a03
        Label1(3).Caption = 0
    Else
        Label1(11).Caption = a03
        Label1(3).Caption = 0
    End If
End If
'----华丽的分割线----
If a03 <> 0 And a07 <> 0 And a11 <> 0 And a15 = 0 Then
    If a03 = a07 And a07 <> a11 Then
        Label1(15).Caption = a11
        Label1(11).Caption = 2 * a03
        S = S + 2 * a03
        Label1(7).Caption = 0
        Label1(3).Caption = 0
    Else
        If a07 = a11 And a03 <> a07 Then
            Label1(15).Caption = 2 * a11
            S = S + 2 * a11
            Label1(11).Caption = a03
            Label1(7).Caption = 0
            Label1(3).Caption = 0
        Else
            If a03 = a07 And a03 = a11 Then
                Label1(15).Caption = 2 * a11
                S = S + 2 * a11
                Label1(11).Caption = a03
                Label1(7).Caption = 0
                Label1(3).Caption = 0
            Else
                Label1(15).Caption = a11
                Label1(11).Caption = a07
                Label1(7).Caption = a03
                Label1(3).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a03 = 0 And a07 <> 0 And a11 <> 0 And a15 <> 0 Then
    If a07 = a11 And a11 <> a15 Then
        Label1(11).Caption = 2 * a07
        S = S + 2 * a07
        Label1(7).Caption = 0
        Label1(3).Caption = 0
    Else
        If a11 = a15 And a07 <> a11 Then
            Label1(15).Caption = 2 * a11
            S = S + 2 * a11
            Label1(11).Caption = a07
            Label1(7).Caption = 0
            Label1(3).Caption = 0
        Else
            If a07 = a11 And a07 = a15 Then
                Label1(15).Caption = 2 * a11
                S = S + 2 * a11
                Label1(11).Caption = a07
                Label1(7).Caption = 0
                Label1(3).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a03 <> 0 And a07 <> 0 And a11 = 0 And a15 <> 0 Then
    If a03 = a07 And a03 <> a15 Then
        Label1(11).Caption = 2 * a07
        S = S + 2 * a07
        Label1(7).Caption = 0
        Label1(3).Caption = 0
    Else
        If a03 <> a07 And a07 = a15 Then
            Label1(15).Caption = 2 * a07
            S = S + 2 * a07
            Label1(11).Caption = a03
            Label1(7).Caption = 0
            Label1(3).Caption = 0
        Else
            If a03 = a07 And a07 = a15 Then
                Label1(15).Caption = 2 * a07
                S = S + 2 * a07
                Label1(11).Caption = a03
                Label1(7).Caption = 0
                Label1(3).Caption = 0
            Else
                Label1(11).Caption = a07
                Label1(7).Caption = a03
                Label1(3).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a03 <> 0 And a07 = 0 And a11 <> 0 And a15 <> 0 Then
    If a03 = a11 And a03 <> a15 Then
        Label1(11).Caption = 2 * a03
        S = S + 2 * a03
        Label1(7).Caption = 0
        Label1(3).Caption = 0
    Else
        If a03 <> a11 And a11 = a15 Then
            Label1(15).Caption = 2 * a11
            S = S + 2 * a11
            Label1(11).Caption = a03
            Label1(7).Caption = 0
            Label1(3).Caption = 0
        Else
            If a03 = a11 And a11 = a15 Then
                Label1(15).Caption = 2 * a11
                S = S + 2 * a11
                Label1(11).Caption = a03
                Label1(7).Caption = 0
                Label1(3).Caption = 0
            Else
                Label1(7).Caption = a03
                Label1(3).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a03 <> 0 And a07 <> 0 And a11 <> 0 And a15 <> 0 Then
    If a03 = a07 And a11 = a15 Then
         Label1(15).Caption = 2 * a11
         S = S + 2 * a11
         Label1(11).Caption = 2 * a03
         S = S + 2 * a03
         Label1(7).Caption = 0
         Label1(3).Caption = 0
    Else
        If a03 <> a07 And a07 = a11 And a11 <> a15 Then
            Label1(11).Caption = 2 * a07
            S = S + 2 * a07
            Label1(7).Caption = a03
            Label1(3).Caption = 0
        Else
            If a03 = a07 And a07 <> a11 Then
                Label1(7).Caption = 2 * a03
                S = S + 2 * a03
                Label1(3).Caption = 0
            Else
                If a07 <> a11 And a11 = a15 Then
                    Label1(15).Caption = 2 * a11
                    S = S + 2 * a11
                    Label1(11).Caption = a07
                    Label1(7).Caption = a03
                    Label1(3).Caption = 0
                Else
                    If a03 = a07 And a07 = a11 And a11 <> a15 Then
                        Label1(11).Caption = 2 * a07
                        S = S + 2 * a07
                        Label1(7).Caption = a03
                        Label1(3).Caption = 0
                    Else
                        If a03 <> a07 And a07 = a11 And a11 = a15 Then
                            Label1(15).Caption = 2 * a11
                            S = S + 2 * a11
                            Label1(11).Caption = a07
                            Label1(7).Caption = a03
                            Label1(3).Caption = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
'--------超级华丽的分割线--------
'--------超级华丽的分割线--------
If a04 <> 0 And a08 = 0 And a12 = 0 And a16 = 0 Then
    Label1(16).Caption = a04
    Label1(4).Caption = 0
End If

If a04 = 0 And a08 <> 0 And a12 = 0 And a16 = 0 Then
    Label1(16).Caption = a08
    Label1(8).Caption = 0
End If

If a04 = 0 And a08 = 0 And a12 <> 0 And a16 = 0 Then
    Label1(16).Caption = a12
    Label1(12).Caption = 0
End If

If a04 = 0 And a08 = 0 And a12 = 0 And a16 <> 0 Then

End If

'----华丽的分割线----

If a04 <> 0 And a08 <> 0 And a12 = 0 And a16 = 0 Then
    If a04 = a08 Then
        Label1(16).Caption = 2 * a04
        S = S + 2 * a04
        Label1(4).Caption = 0
        Label1(8).Caption = 0
    Else
        Label1(12).Caption = a04
        Label1(16).Caption = a08
        Label1(4).Caption = 0
        Label1(8).Caption = 0
    End If
End If

If a04 = 0 And a08 <> 0 And a12 <> 0 And a16 = 0 Then
    If a08 = a12 Then
        Label1(16).Caption = 2 * a08
        S = S + 2 * a08
        Label1(8).Caption = 0
        Label1(12).Caption = 0
    Else
        Label1(16).Caption = a12
        Label1(12).Caption = a08
        Label1(8).Caption = 0
    End If
End If

If a04 = 0 And a08 = 0 And a12 <> 0 And a16 <> 0 Then
    If a12 = a16 Then
        Label1(16).Caption = 2 * a12
        S = S + 2 * a12
        Label1(12).Caption = 0
    End If
End If

If a04 <> 0 And a08 = 0 And a12 <> 0 And a16 = 0 Then
    If a04 = a12 Then
        Label1(16).Caption = 2 * a04
        S = S + 2 * a04
        Label1(4).Caption = 0
        Label1(12).Caption = 0
    Else
        Label1(16).Caption = a12
        Label1(12).Caption = a04
        Label1(4).Caption = 0
    End If
End If

If a04 = 0 And a08 <> 0 And a12 = 0 And a16 <> 0 Then
    If a08 = a16 Then
        Label1(16).Caption = 2 * a08
        S = S + 2 * a08
        Label1(8).Caption = 0
    Else
        Label1(12).Caption = a08
        Label1(8).Caption = 0
    End If
End If

If a04 <> 0 And a08 = 0 And a12 = 0 And a16 <> 0 Then
    If a04 = a16 Then
        Label1(16).Caption = 2 * a04
        S = S + 2 * a04
        Label1(4).Caption = 0
    Else
        Label1(12).Caption = a04
        Label1(4).Caption = 0
    End If
End If
'----华丽的分割线----
If a04 <> 0 And a08 <> 0 And a12 <> 0 And a16 = 0 Then
    If a04 = a08 And a08 <> a12 Then
        Label1(16).Caption = a12
        Label1(12).Caption = 2 * a04
        S = S + 2 * a04
        Label1(8).Caption = 0
        Label1(4).Caption = 0
    Else
        If a08 = a12 And a04 <> a08 Then
            Label1(16).Caption = 2 * a12
            S = S + 2 * a12
            Label1(12).Caption = a04
            Label1(8).Caption = 0
            Label1(4).Caption = 0
        Else
            If a04 = a08 And a04 = a12 Then
                Label1(16).Caption = 2 * a12
                S = S + 2 * a12
                Label1(12).Caption = a04
                Label1(8).Caption = 0
                Label1(4).Caption = 0
            Else
                Label1(16).Caption = a12
                Label1(12).Caption = a08
                Label1(8).Caption = a04
                Label1(4).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a04 = 0 And a08 <> 0 And a12 <> 0 And a16 <> 0 Then
    If a08 = a12 And a12 <> a16 Then
        Label1(12).Caption = 2 * a08
        S = S + 2 * a08
        Label1(8).Caption = 0
        Label1(4).Caption = 0
    Else
        If a12 = a16 And a08 <> a12 Then
            Label1(16).Caption = 2 * a12
            S = S + 2 * a12
            Label1(12).Caption = a08
            Label1(8).Caption = 0
            Label1(4).Caption = 0
        Else
            If a08 = a12 And a08 = a16 Then
                Label1(16).Caption = 2 * a12
                S = S + 2 * a12
                Label1(12).Caption = a08
                Label1(8).Caption = 0
                Label1(4).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a04 <> 0 And a08 <> 0 And a12 = 0 And a16 <> 0 Then
    If a04 = a08 And a04 <> a16 Then
        Label1(12).Caption = 2 * a08
        S = S + 2 * a08
        Label1(8).Caption = 0
        Label1(4).Caption = 0
    Else
        If a04 <> a08 And a08 = a16 Then
            Label1(16).Caption = 2 * a08
            S = S + 2 * a08
            Label1(12).Caption = a04
            Label1(8).Caption = 0
            Label1(4).Caption = 0
        Else
            If a04 = a08 And a08 = a16 Then
                Label1(16).Caption = 2 * a08
                S = S + 2 * a08
                Label1(12).Caption = a04
                Label1(8).Caption = 0
                Label1(4).Caption = 0
            Else
                Label1(12).Caption = a08
                Label1(8).Caption = a04
                Label1(4).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a04 <> 0 And a08 = 0 And a12 <> 0 And a16 <> 0 Then
    If a04 = a12 And a04 <> a16 Then
        Label1(12).Caption = 2 * a04
        S = S + 2 * a04
        Label1(8).Caption = 0
        Label1(4).Caption = 0
    Else
        If a04 <> a12 And a12 = a16 Then
            Label1(16).Caption = 2 * a12
            S = S + 2 * a12
            Label1(12).Caption = a04
            Label1(8).Caption = 0
            Label1(4).Caption = 0
        Else
            If a04 = a12 And a12 = a16 Then
                Label1(16).Caption = 2 * a12
                S = S + 2 * a12
                Label1(12).Caption = a04
                Label1(8).Caption = 0
                Label1(4).Caption = 0
            Else
                Label1(8).Caption = a04
                Label1(4).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a04 <> 0 And a08 <> 0 And a12 <> 0 And a16 <> 0 Then
    If a04 = a08 And a12 = a16 Then
         Label1(16).Caption = 2 * a12
         S = S + 2 * a12
         Label1(12).Caption = 2 * a04
         S = S + 2 * a04
         Label1(8).Caption = 0
         Label1(4).Caption = 0
    Else
        If a04 <> a08 And a08 = a12 And a12 <> a16 Then
            Label1(12).Caption = 2 * a08
            S = S + 2 * a08
            Label1(8).Caption = a04
            Label1(4).Caption = 0
        Else
            If a04 = a08 And a08 <> a12 Then
                Label1(8).Caption = 2 * a04
                S = S + 2 * a04
                Label1(4).Caption = 0
            Else
                If a08 <> a12 And a12 = a16 Then
                    Label1(16).Caption = 2 * a12
                    S = S + 2 * a12
                    Label1(12).Caption = a08
                    Label1(8).Caption = a04
                    Label1(4).Caption = 0
                Else
                    If a04 = a08 And a08 = a12 And a12 <> a16 Then
                        Label1(12).Caption = 2 * a08
                        S = S + 2 * a08
                        Label1(8).Caption = a04
                        Label1(4).Caption = 0
                    Else
                        If a04 <> a08 And a08 = a12 And a12 = a16 Then
                            Label1(16).Caption = 2 * a12
                            S = S + 2 * a12
                            Label1(12).Caption = a08
                            Label1(8).Caption = a04
                            Label1(4).Caption = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
Call Command101_Click
Call Command5_Click
Call Command6_Click
End Sub

Private Sub Command3_Click()
Call Command66_Click

If a04 <> 0 And a03 = 0 And a02 = 0 And a01 = 0 Then
    Label1(1).Caption = a04
    Label1(4).Caption = 0
End If

If a04 = 0 And a03 <> 0 And a02 = 0 And a01 = 0 Then
    Label1(1).Caption = a03
    Label1(3).Caption = 0
End If

If a04 = 0 And a03 = 0 And a02 <> 0 And a01 = 0 Then
    Label1(1).Caption = a02
    Label1(2).Caption = 0
End If

If a04 = 0 And a03 = 0 And a02 = 0 And a01 <> 0 Then

End If

'----华丽的分割线----

If a04 <> 0 And a03 <> 0 And a02 = 0 And a01 = 0 Then
    If a04 = a03 Then
        Label1(1).Caption = 2 * a04
        S = S + 2 * a04
        Label1(4).Caption = 0
        Label1(3).Caption = 0
    Else
        Label1(2).Caption = a04
        Label1(1).Caption = a03
        Label1(4).Caption = 0
        Label1(3).Caption = 0
    End If
End If

If a04 = 0 And a03 <> 0 And a02 <> 0 And a01 = 0 Then
    If a03 = a02 Then
        Label1(1).Caption = 2 * a03
        S = S + 2 * a03
        Label1(3).Caption = 0
        Label1(2).Caption = 0
    Else
        Label1(1).Caption = a02
        Label1(2).Caption = a03
        Label1(3).Caption = 0
    End If
End If

If a04 = 0 And a03 = 0 And a02 <> 0 And a01 <> 0 Then
    If a02 = a01 Then
        Label1(1).Caption = 2 * a02
        S = S + 2 * a02
        Label1(2).Caption = 0
    End If
End If

If a04 <> 0 And a03 = 0 And a02 <> 0 And a01 = 0 Then
    If a04 = a02 Then
        Label1(1).Caption = 2 * a04
        S = S + 2 * a04
        Label1(4).Caption = 0
        Label1(2).Caption = 0
    Else
        Label1(1).Caption = a02
        Label1(2).Caption = a04
        Label1(4).Caption = 0
    End If
End If

If a04 = 0 And a03 <> 0 And a02 = 0 And a01 <> 0 Then
    If a03 = a01 Then
        Label1(1).Caption = 2 * a03
        S = S + 2 * a03
        Label1(3).Caption = 0
    Else
        Label1(2).Caption = a03
        Label1(3).Caption = 0
    End If
End If

If a04 <> 0 And a03 = 0 And a02 = 0 And a01 <> 0 Then
    If a04 = a01 Then
        Label1(1).Caption = 2 * a04
        S = S + 2 * a04
        Label1(4).Caption = 0
    Else
        Label1(2).Caption = a04
        Label1(4).Caption = 0
    End If
End If
'----华丽的分割线----
If a04 <> 0 And a03 <> 0 And a02 <> 0 And a01 = 0 Then
    If a04 = a03 And a03 <> a02 Then
        Label1(1).Caption = a02
        Label1(2).Caption = 2 * a04
        S = S + 2 * a04
        Label1(3).Caption = 0
        Label1(4).Caption = 0
    Else
        If a03 = a02 And a04 <> a03 Then
            Label1(1).Caption = 2 * a02
            S = S + 2 * a02
            Label1(2).Caption = a04
            Label1(3).Caption = 0
            Label1(4).Caption = 0
        Else
            If a04 = a03 And a04 = a02 Then
                Label1(1).Caption = 2 * a02
                S = S + 2 * a02
                Label1(2).Caption = a04
                Label1(3).Caption = 0
                Label1(4).Caption = 0
            Else
                Label1(1).Caption = a02
                Label1(2).Caption = a03
                Label1(3).Caption = a04
                Label1(4).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a04 = 0 And a03 <> 0 And a02 <> 0 And a01 <> 0 Then
    If a03 = a02 And a02 <> a01 Then
        Label1(2).Caption = 2 * a03
        S = S + 2 * a03
        Label1(3).Caption = 0
        Label1(4).Caption = 0
    Else
        If a02 = a01 And a03 <> a02 Then
            Label1(1).Caption = 2 * a02
            S = S + 2 * a02
            Label1(2).Caption = a03
            Label1(3).Caption = 0
            Label1(4).Caption = 0
        Else
            If a03 = a02 And a03 = a01 Then
                Label1(1).Caption = 2 * a02
                S = S + 2 * a02
                Label1(2).Caption = a03
                Label1(3).Caption = 0
                Label1(4).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a04 <> 0 And a03 <> 0 And a02 = 0 And a01 <> 0 Then
    If a04 = a03 And a04 <> a01 Then
        Label1(2).Caption = 2 * a03
        S = S + 2 * a03
        Label1(3).Caption = 0
        Label1(4).Caption = 0
    Else
        If a04 <> a03 And a03 = a01 Then
            Label1(1).Caption = 2 * a03
            S = S + 2 * a03
            Label1(2).Caption = a04
            Label1(3).Caption = 0
            Label1(4).Caption = 0
        Else
            If a04 = a03 And a03 = a01 Then
                Label1(1).Caption = 2 * a03
                S = S + 2 * a03
                Label1(2).Caption = a04
                Label1(3).Caption = 0
                Label1(4).Caption = 0
            Else
                Label1(2).Caption = a03
                Label1(3).Caption = a04
                Label1(4).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a04 <> 0 And a03 = 0 And a02 <> 0 And a01 <> 0 Then
    If a04 = a02 And a04 <> a01 Then
        Label1(2).Caption = 2 * a04
        S = S + 2 * a04
        Label1(3).Caption = 0
        Label1(4).Caption = 0
    Else
        If a04 <> a02 And a02 = a01 Then
            Label1(1).Caption = 2 * a02
            S = S + 2 * a02
            Label1(2).Caption = a04
            Label1(3).Caption = 0
            Label1(4).Caption = 0
        Else
            If a04 = a02 And a02 = a01 Then
                Label1(1).Caption = 2 * a02
                S = S + 2 * a02
                Label1(2).Caption = a04
                Label1(3).Caption = 0
                Label1(4).Caption = 0
            Else
                Label1(3).Caption = a04
                Label1(4).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a04 <> 0 And a03 <> 0 And a02 <> 0 And a01 <> 0 Then
    If a04 = a03 And a02 = a01 Then
         Label1(1).Caption = 2 * a02
         S = S + 2 * a02
         Label1(2).Caption = 2 * a04
         S = S + 2 * a04
         Label1(3).Caption = 0
         Label1(4).Caption = 0
    Else
        If a04 <> a03 And a03 = a02 And a02 <> a01 Then
            Label1(2).Caption = 2 * a03
            S = S + 2 * a03
            Label1(3).Caption = a04
            Label1(4).Caption = 0
        Else
            If a04 = a03 And a03 <> a02 Then
                Label1(3).Caption = 2 * a04
                S = S + 2 * a04
                Label1(4).Caption = 0
            Else
                If a03 <> a02 And a02 = a01 Then
                    Label1(1).Caption = 2 * a02
                    S = S + 2 * a02
                    Label1(2).Caption = a03
                    Label1(3).Caption = a04
                    Label1(4).Caption = 0
                Else
                    If a04 = a03 And a03 = a02 And a02 <> a01 Then
                        Label1(2).Caption = 2 * a03
                        S = S + 2 * a03
                        Label1(3).Caption = a04
                        Label1(4).Caption = 0
                    Else
                        If a04 <> a03 And a03 = a02 And a02 = a01 Then
                            Label1(1).Caption = 2 * a02
                            S = S + 2 * a02
                            Label1(2).Caption = a03
                            Label1(3).Caption = a04
                            Label1(4).Caption = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
'--------超级华丽的分割线--------
'--------超级华丽的分割线--------
If a08 <> 0 And a07 = 0 And a06 = 0 And a05 = 0 Then
    Label1(5).Caption = a08
    Label1(8).Caption = 0
End If

If a08 = 0 And a07 <> 0 And a06 = 0 And a05 = 0 Then
    Label1(5).Caption = a07
    Label1(7).Caption = 0
End If

If a08 = 0 And a07 = 0 And a06 <> 0 And a05 = 0 Then
    Label1(5).Caption = a06
    Label1(6).Caption = 0
End If

If a08 = 0 And a07 = 0 And a06 = 0 And a05 <> 0 Then

End If

'----华丽的分割线----

If a08 <> 0 And a07 <> 0 And a06 = 0 And a05 = 0 Then
    If a08 = a07 Then
        Label1(5).Caption = 2 * a08
        S = S + 2 * a08
        Label1(8).Caption = 0
        Label1(7).Caption = 0
    Else
        Label1(6).Caption = a08
        Label1(5).Caption = a07
        Label1(8).Caption = 0
        Label1(7).Caption = 0
    End If
End If

If a08 = 0 And a07 <> 0 And a06 <> 0 And a05 = 0 Then
    If a07 = a06 Then
        Label1(5).Caption = 2 * a07
        S = S + 2 * a07
        Label1(7).Caption = 0
        Label1(6).Caption = 0
    Else
        Label1(5).Caption = a06
        Label1(6).Caption = a07
        Label1(7).Caption = 0
    End If
End If

If a08 = 0 And a07 = 0 And a06 <> 0 And a05 <> 0 Then
    If a06 = a05 Then
        Label1(5).Caption = 2 * a06
        S = S + 2 * a06
        Label1(6).Caption = 0
    End If
End If

If a08 <> 0 And a07 = 0 And a06 <> 0 And a05 = 0 Then
    If a08 = a06 Then
        Label1(5).Caption = 2 * a08
        S = S + 2 * a08
        Label1(8).Caption = 0
        Label1(6).Caption = 0
    Else
        Label1(5).Caption = a06
        Label1(6).Caption = a08
        Label1(8).Caption = 0
    End If
End If

If a08 = 0 And a07 <> 0 And a06 = 0 And a05 <> 0 Then
    If a07 = a05 Then
        Label1(5).Caption = 2 * a07
        S = S + 2 * a07
        Label1(7).Caption = 0
    Else
        Label1(6).Caption = a07
        Label1(7).Caption = 0
    End If
End If

If a08 <> 0 And a07 = 0 And a06 = 0 And a05 <> 0 Then
    If a08 = a05 Then
        Label1(5).Caption = 2 * a08
        S = S + 2 * a08
        Label1(8).Caption = 0
    Else
        Label1(6).Caption = a08
        Label1(8).Caption = 0
    End If
End If
'----华丽的分割线----
If a08 <> 0 And a07 <> 0 And a06 <> 0 And a05 = 0 Then
    If a08 = a07 And a07 <> a06 Then
        Label1(5).Caption = a06
        Label1(6).Caption = 2 * a08
        S = S + 2 * a08
        Label1(7).Caption = 0
        Label1(8).Caption = 0
    Else
        If a07 = a06 And a08 <> a07 Then
            Label1(5).Caption = 2 * a06
            S = S + 2 * a06
            Label1(6).Caption = a08
            Label1(7).Caption = 0
            Label1(8).Caption = 0
        Else
            If a08 = a07 And a08 = a06 Then
                Label1(5).Caption = 2 * a06
                S = S + 2 * a06
                Label1(6).Caption = a08
                Label1(7).Caption = 0
                Label1(8).Caption = 0
            Else
                Label1(5).Caption = a06
                Label1(6).Caption = a07
                Label1(7).Caption = a08
                Label1(8).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a08 = 0 And a07 <> 0 And a06 <> 0 And a05 <> 0 Then
    If a07 = a06 And a06 <> a05 Then
        Label1(6).Caption = 2 * a07
        S = S + 2 * a07
        Label1(7).Caption = 0
        Label1(8).Caption = 0
    Else
        If a06 = a05 And a07 <> a06 Then
            Label1(5).Caption = 2 * a06
            S = S + 2 * a06
            Label1(6).Caption = a07
            Label1(7).Caption = 0
            Label1(8).Caption = 0
        Else
            If a07 = a06 And a07 = a05 Then
                Label1(5).Caption = 2 * a06
                S = S + 2 * a06
                Label1(6).Caption = a07
                Label1(7).Caption = 0
                Label1(8).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a08 <> 0 And a07 <> 0 And a06 = 0 And a05 <> 0 Then
    If a08 = a07 And a08 <> a05 Then
        Label1(6).Caption = 2 * a07
        S = S + 2 * a07
        Label1(7).Caption = 0
        Label1(8).Caption = 0
    Else
        If a08 <> a07 And a07 = a05 Then
            Label1(5).Caption = 2 * a07
            S = S + 2 * a07
            Label1(6).Caption = a08
            Label1(7).Caption = 0
            Label1(8).Caption = 0
        Else
            If a08 = a07 And a07 = a05 Then
                Label1(5).Caption = 2 * a07
                S = S + 2 * a07
                Label1(6).Caption = a08
                Label1(7).Caption = 0
                Label1(8).Caption = 0
            Else
                Label1(6).Caption = a07
                Label1(7).Caption = a08
                Label1(8).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a08 <> 0 And a07 = 0 And a06 <> 0 And a05 <> 0 Then
    If a08 = a06 And a08 <> a05 Then
        Label1(6).Caption = 2 * a08
        S = S + 2 * a08
        Label1(7).Caption = 0
        Label1(8).Caption = 0
    Else
        If a08 <> a06 And a06 = a05 Then
            Label1(5).Caption = 2 * a06
            S = S + 2 * a06
            Label1(6).Caption = a08
            Label1(7).Caption = 0
            Label1(8).Caption = 0
        Else
            If a08 = a06 And a06 = a05 Then
                Label1(5).Caption = 2 * a06
                S = S + 2 * a06
                Label1(6).Caption = a08
                Label1(7).Caption = 0
                Label1(8).Caption = 0
            Else
                Label1(7).Caption = a08
                Label1(8).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a08 <> 0 And a07 <> 0 And a06 <> 0 And a05 <> 0 Then
    If a08 = a07 And a06 = a05 Then
         Label1(5).Caption = 2 * a06
         S = S + 2 * a06
         Label1(6).Caption = 2 * a08
         S = S + 2 * a08
         Label1(7).Caption = 0
         Label1(8).Caption = 0
    Else
        If a08 <> a07 And a07 = a06 And a06 <> a05 Then
            Label1(6).Caption = 2 * a07
            S = S + 2 * a07
            Label1(7).Caption = a08
            Label1(8).Caption = 0
        Else
            If a08 = a07 And a07 <> a06 Then
                Label1(7).Caption = 2 * a08
                S = S + 2 * a08
                Label1(8).Caption = 0
            Else
                If a07 <> a06 And a06 = a05 Then
                    Label1(5).Caption = 2 * a06
                    S = S + 2 * a06
                    Label1(6).Caption = a07
                    Label1(7).Caption = a08
                    Label1(8).Caption = 0
                Else
                    If a08 = a07 And a07 = a06 And a06 <> a05 Then
                        Label1(6).Caption = 2 * a07
                        S = S + 2 * a07
                        Label1(7).Caption = a08
                        Label1(8).Caption = 0
                    Else
                        If a08 <> a07 And a07 = a06 And a06 = a05 Then
                            Label1(5).Caption = 2 * a06
                            S = S + 2 * a06
                            Label1(6).Caption = a07
                            Label1(7).Caption = a08
                            Label1(8).Caption = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
'--------超级华丽的分割线--------
'--------超级华丽的分割线--------
Call Command66_Click

If a12 <> 0 And a11 = 0 And a10 = 0 And a09 = 0 Then
    Label1(9).Caption = a12
    Label1(12).Caption = 0
End If

If a12 = 0 And a11 <> 0 And a10 = 0 And a09 = 0 Then
    Label1(9).Caption = a11
    Label1(11).Caption = 0
End If

If a12 = 0 And a11 = 0 And a10 <> 0 And a09 = 0 Then
    Label1(9).Caption = a10
    Label1(10).Caption = 0
End If

If a12 = 0 And a11 = 0 And a10 = 0 And a09 <> 0 Then

End If

'----华丽的分割线----

If a12 <> 0 And a11 <> 0 And a10 = 0 And a09 = 0 Then
    If a12 = a11 Then
        Label1(9).Caption = 2 * a12
        S = S + 2 * a12
        Label1(12).Caption = 0
        Label1(11).Caption = 0
    Else
        Label1(10).Caption = a12
        Label1(9).Caption = a11
        Label1(12).Caption = 0
        Label1(11).Caption = 0
    End If
End If

If a12 = 0 And a11 <> 0 And a10 <> 0 And a09 = 0 Then
    If a11 = a10 Then
        Label1(9).Caption = 2 * a11
        S = S + 2 * a11
        Label1(11).Caption = 0
        Label1(10).Caption = 0
    Else
        Label1(9).Caption = a10
        Label1(10).Caption = a11
        Label1(11).Caption = 0
    End If
End If

If a12 = 0 And a11 = 0 And a10 <> 0 And a09 <> 0 Then
    If a10 = a09 Then
        Label1(9).Caption = 2 * a10
        S = S + 2 * a10
        Label1(10).Caption = 0
    End If
End If

If a12 <> 0 And a11 = 0 And a10 <> 0 And a09 = 0 Then
    If a12 = a10 Then
        Label1(9).Caption = 2 * a12
        S = S + 2 * a12
        Label1(12).Caption = 0
        Label1(10).Caption = 0
    Else
        Label1(9).Caption = a10
        Label1(10).Caption = a12
        Label1(12).Caption = 0
    End If
End If

If a12 = 0 And a11 <> 0 And a10 = 0 And a09 <> 0 Then
    If a11 = a09 Then
        Label1(9).Caption = 2 * a11
        S = S + 2 * a11
        Label1(11).Caption = 0
    Else
        Label1(10).Caption = a11
        Label1(11).Caption = 0
    End If
End If

If a12 <> 0 And a11 = 0 And a10 = 0 And a09 <> 0 Then
    If a12 = a09 Then
        Label1(9).Caption = 2 * a12
        S = S + 2 * a12
        Label1(12).Caption = 0
    Else
        Label1(10).Caption = a12
        Label1(12).Caption = 0
    End If
End If
'----华丽的分割线----
If a12 <> 0 And a11 <> 0 And a10 <> 0 And a09 = 0 Then
    If a12 = a11 And a11 <> a10 Then
        Label1(9).Caption = a10
        Label1(10).Caption = 2 * a12
        S = S + 2 * a12
        Label1(11).Caption = 0
        Label1(12).Caption = 0
    Else
        If a11 = a10 And a12 <> a11 Then
            Label1(9).Caption = 2 * a10
            S = S + 2 * a10
            Label1(10).Caption = a12
            Label1(11).Caption = 0
            Label1(12).Caption = 0
        Else
            If a12 = a11 And a12 = a10 Then
                Label1(9).Caption = 2 * a10
                S = S + 2 * a10
                Label1(10).Caption = a12
                Label1(11).Caption = 0
                Label1(12).Caption = 0
            Else
                Label1(9).Caption = a10
                Label1(10).Caption = a11
                Label1(11).Caption = a12
                Label1(12).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a12 = 0 And a11 <> 0 And a10 <> 0 And a09 <> 0 Then
    If a11 = a10 And a10 <> a09 Then
        Label1(10).Caption = 2 * a11
        S = S + 2 * a11
        Label1(11).Caption = 0
        Label1(12).Caption = 0
    Else
        If a10 = a09 And a11 <> a10 Then
            Label1(9).Caption = 2 * a10
            S = S + 2 * a10
            Label1(10).Caption = a11
            Label1(11).Caption = 0
            Label1(12).Caption = 0
        Else
            If a11 = a10 And a11 = a09 Then
                Label1(9).Caption = 2 * a10
                S = S + 2 * a10
                Label1(10).Caption = a11
                Label1(11).Caption = 0
                Label1(12).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a12 <> 0 And a11 <> 0 And a10 = 0 And a09 <> 0 Then
    If a12 = a11 And a12 <> a09 Then
        Label1(10).Caption = 2 * a11
        S = S + 2 * a11
        Label1(11).Caption = 0
        Label1(12).Caption = 0
    Else
        If a12 <> a11 And a11 = a09 Then
            Label1(9).Caption = 2 * a11
            S = S + 2 * a11
            Label1(10).Caption = a12
            Label1(11).Caption = 0
            Label1(12).Caption = 0
        Else
            If a12 = a11 And a11 = a09 Then
                Label1(9).Caption = 2 * a11
                S = S + 2 * a11
                Label1(10).Caption = a12
                Label1(11).Caption = 0
                Label1(12).Caption = 0
            Else
                Label1(10).Caption = a11
                Label1(11).Caption = a12
                Label1(12).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a12 <> 0 And a11 = 0 And a10 <> 0 And a09 <> 0 Then
    If a12 = a10 And a12 <> a09 Then
        Label1(10).Caption = 2 * a12
        S = S + 2 * a12
        Label1(11).Caption = 0
        Label1(12).Caption = 0
    Else
        If a12 <> a10 And a10 = a09 Then
            Label1(9).Caption = 2 * a10
            S = S + 2 * a10
            Label1(10).Caption = a12
            Label1(11).Caption = 0
            Label1(12).Caption = 0
        Else
            If a12 = a10 And a10 = a09 Then
                Label1(9).Caption = 2 * a10
                S = S + 2 * a10
                Label1(10).Caption = a12
                Label1(11).Caption = 0
                Label1(12).Caption = 0
            Else
                Label1(11).Caption = a12
                Label1(12).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a12 <> 0 And a11 <> 0 And a10 <> 0 And a09 <> 0 Then
    If a12 = a11 And a10 = a09 Then
         Label1(9).Caption = 2 * a10
         S = S + 2 * a10
         Label1(10).Caption = 2 * a12
         S = S + 2 * a12
         Label1(11).Caption = 0
         Label1(12).Caption = 0
    Else
        If a12 <> a11 And a11 = a10 And a10 <> a09 Then
            Label1(10).Caption = 2 * a11
            S = S + 2 * a11
            Label1(11).Caption = a12
            Label1(12).Caption = 0
        Else
            If a12 = a11 And a11 <> a10 Then
                Label1(11).Caption = 2 * a12
                S = S + 2 * a12
                Label1(12).Caption = 0
            Else
                If a11 <> a10 And a10 = a09 Then
                    Label1(9).Caption = 2 * a10
                    S = S + 2 * a10
                    Label1(10).Caption = a11
                    Label1(11).Caption = a12
                    Label1(12).Caption = 0
                Else
                    If a12 = a11 And a11 = a10 And a10 <> a09 Then
                        Label1(10).Caption = 2 * a11
                        S = S + 2 * a11
                        Label1(11).Caption = a12
                        Label1(12).Caption = 0
                    Else
                        If a12 <> a11 And a11 = a10 And a10 = a09 Then
                            Label1(9).Caption = 2 * a10
                            S = S + 2 * a10
                            Label1(10).Caption = a11
                            Label1(11).Caption = a12
                            Label1(12).Caption = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
'--------超级华丽的分割线--------
'--------超级华丽的分割线--------
If a16 <> 0 And a15 = 0 And a14 = 0 And a13 = 0 Then
    Label1(13).Caption = a16
    Label1(16).Caption = 0
End If

If a16 = 0 And a15 <> 0 And a14 = 0 And a13 = 0 Then
    Label1(13).Caption = a15
    Label1(15).Caption = 0
End If

If a16 = 0 And a15 = 0 And a14 <> 0 And a13 = 0 Then
    Label1(13).Caption = a14
    Label1(14).Caption = 0
End If

If a16 = 0 And a15 = 0 And a14 = 0 And a13 <> 0 Then

End If

'----华丽的分割线----

If a16 <> 0 And a15 <> 0 And a14 = 0 And a13 = 0 Then
    If a16 = a15 Then
        Label1(13).Caption = 2 * a16
        S = S + 2 * a16
        Label1(16).Caption = 0
        Label1(15).Caption = 0
    Else
        Label1(14).Caption = a16
        Label1(13).Caption = a15
        Label1(16).Caption = 0
        Label1(15).Caption = 0
    End If
End If

If a16 = 0 And a15 <> 0 And a14 <> 0 And a13 = 0 Then
    If a15 = a14 Then
        Label1(13).Caption = 2 * a15
        S = S + 2 * a15
        Label1(15).Caption = 0
        Label1(14).Caption = 0
    Else
        Label1(13).Caption = a14
        Label1(14).Caption = a15
        Label1(15).Caption = 0
    End If
End If

If a16 = 0 And a15 = 0 And a14 <> 0 And a13 <> 0 Then
    If a14 = a13 Then
        Label1(13).Caption = 2 * a14
        S = S + 2 * a14
        Label1(14).Caption = 0
    End If
End If

If a16 <> 0 And a15 = 0 And a14 <> 0 And a13 = 0 Then
    If a16 = a14 Then
        Label1(13).Caption = 2 * a16
        S = S + 2 * a16
        Label1(16).Caption = 0
        Label1(14).Caption = 0
    Else
        Label1(13).Caption = a14
        Label1(14).Caption = a16
        Label1(16).Caption = 0
    End If
End If

If a16 = 0 And a15 <> 0 And a14 = 0 And a13 <> 0 Then
    If a15 = a13 Then
        Label1(13).Caption = 2 * a15
        S = S + 2 * a15
        Label1(15).Caption = 0
    Else
        Label1(14).Caption = a15
        Label1(15).Caption = 0
    End If
End If

If a16 <> 0 And a15 = 0 And a14 = 0 And a13 <> 0 Then
    If a16 = a13 Then
        Label1(13).Caption = 2 * a16
        S = S + 2 * a16
        Label1(16).Caption = 0
    Else
        Label1(14).Caption = a16
        Label1(16).Caption = 0
    End If
End If
'----华丽的分割线----
If a16 <> 0 And a15 <> 0 And a14 <> 0 And a13 = 0 Then
    If a16 = a15 And a15 <> a14 Then
        Label1(13).Caption = a14
        Label1(14).Caption = 2 * a16
        S = S + 2 * a16
        Label1(15).Caption = 0
        Label1(16).Caption = 0
    Else
        If a15 = a14 And a16 <> a15 Then
            Label1(13).Caption = 2 * a14
            S = S + 2 * a14
            Label1(14).Caption = a16
            Label1(15).Caption = 0
            Label1(16).Caption = 0
        Else
            If a16 = a15 And a16 = a14 Then
                Label1(13).Caption = 2 * a14
                S = S + 2 * a14
                Label1(14).Caption = a16
                Label1(15).Caption = 0
                Label1(16).Caption = 0
            Else
                Label1(13).Caption = a14
                Label1(14).Caption = a15
                Label1(15).Caption = a16
                Label1(16).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a16 = 0 And a15 <> 0 And a14 <> 0 And a13 <> 0 Then
    If a15 = a14 And a14 <> a13 Then
        Label1(14).Caption = 2 * a15
        S = S + 2 * a15
        Label1(15).Caption = 0
        Label1(16).Caption = 0
    Else
        If a14 = a13 And a15 <> a14 Then
            Label1(13).Caption = 2 * a14
            S = S + 2 * a14
            Label1(14).Caption = a15
            Label1(15).Caption = 0
            Label1(16).Caption = 0
        Else
            If a15 = a14 And a15 = a13 Then
                Label1(13).Caption = 2 * a14
                S = S + 2 * a14
                Label1(14).Caption = a15
                Label1(15).Caption = 0
                Label1(16).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a16 <> 0 And a15 <> 0 And a14 = 0 And a13 <> 0 Then
    If a16 = a15 And a16 <> a13 Then
        Label1(14).Caption = 2 * a15
        S = S + 2 * a15
        Label1(15).Caption = 0
        Label1(16).Caption = 0
    Else
        If a16 <> a15 And a15 = a13 Then
            Label1(13).Caption = 2 * a15
            S = S + 2 * a15
            Label1(14).Caption = a16
            Label1(15).Caption = 0
            Label1(16).Caption = 0
        Else
            If a16 = a15 And a15 = a13 Then
                Label1(13).Caption = 2 * a15
                S = S + 2 * a15
                Label1(14).Caption = a16
                Label1(15).Caption = 0
                Label1(16).Caption = 0
            Else
                Label1(14).Caption = a15
                Label1(15).Caption = a16
                Label1(16).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a16 <> 0 And a15 = 0 And a14 <> 0 And a13 <> 0 Then
    If a16 = a14 And a16 <> a13 Then
        Label1(14).Caption = 2 * a16
        S = S + 2 * a16
        Label1(15).Caption = 0
        Label1(16).Caption = 0
    Else
        If a16 <> a14 And a14 = a13 Then
            Label1(13).Caption = 2 * a14
            S = S + 2 * a14
            Label1(14).Caption = a16
            Label1(15).Caption = 0
            Label1(16).Caption = 0
        Else
            If a16 = a14 And a14 = a13 Then
                Label1(13).Caption = 2 * a14
                S = S + 2 * a14
                Label1(14).Caption = a16
                Label1(15).Caption = 0
                Label1(16).Caption = 0
            Else
                Label1(15).Caption = a16
                Label1(16).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a16 <> 0 And a15 <> 0 And a14 <> 0 And a13 <> 0 Then
    If a16 = a15 And a14 = a13 Then
         Label1(13).Caption = 2 * a14
         S = S + 2 * a14
         Label1(14).Caption = 2 * a16
         S = S + 2 * a16
         Label1(15).Caption = 0
         Label1(16).Caption = 0
    Else
        If a16 <> a15 And a15 = a14 And a14 <> a13 Then
            Label1(14).Caption = 2 * a15
            S = S + 2 * a15
            Label1(15).Caption = a16
            Label1(16).Caption = 0
        Else
            If a16 = a15 And a15 <> a14 Then
                Label1(15).Caption = 2 * a16
                S = S + 2 * a16
                Label1(16).Caption = 0
            Else
                If a15 <> a14 And a14 = a13 Then
                    Label1(13).Caption = 2 * a14
                    S = S + 2 * a14
                    Label1(14).Caption = a15
                    Label1(15).Caption = a16
                    Label1(16).Caption = 0
                Else
                    If a16 = a15 And a15 = a14 And a14 <> a13 Then
                        Label1(14).Caption = 2 * a15
                        S = S + 2 * a15
                        Label1(15).Caption = a16
                        Label1(16).Caption = 0
                    Else
                        If a16 <> a15 And a15 = a14 And a14 = a13 Then
                            Label1(13).Caption = 2 * a14
                            S = S + 2 * a14
                            Label1(14).Caption = a15
                            Label1(15).Caption = a16
                            Label1(16).Caption = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
Call Command101_Click
Call Command5_Click
Call Command6_Click
End Sub

Private Sub Command4_Click()
Call Command66_Click

If a13 <> 0 And a09 = 0 And a05 = 0 And a01 = 0 Then
    Label1(1).Caption = a13
    Label1(13).Caption = 0
End If

If a13 = 0 And a09 <> 0 And a05 = 0 And a01 = 0 Then
    Label1(1).Caption = a09
    Label1(9).Caption = 0
End If

If a13 = 0 And a09 = 0 And a05 <> 0 And a01 = 0 Then
    Label1(1).Caption = a05
    Label1(5).Caption = 0
End If

If a13 = 0 And a09 = 0 And a05 = 0 And a01 <> 0 Then

End If

'----华丽的分割线----

If a13 <> 0 And a09 <> 0 And a05 = 0 And a01 = 0 Then
    If a13 = a09 Then
        Label1(1).Caption = 2 * a13
        S = S + 2 * a13
        Label1(13).Caption = 0
        Label1(9).Caption = 0
    Else
        Label1(5).Caption = a13
        Label1(1).Caption = a09
        Label1(13).Caption = 0
        Label1(9).Caption = 0
    End If
End If

If a13 = 0 And a09 <> 0 And a05 <> 0 And a01 = 0 Then
    If a09 = a05 Then
        Label1(1).Caption = 2 * a09
        S = S + 2 * a09
        Label1(9).Caption = 0
        Label1(5).Caption = 0
    Else
        Label1(1).Caption = a05
        Label1(5).Caption = a09
        Label1(9).Caption = 0
    End If
End If

If a13 = 0 And a09 = 0 And a05 <> 0 And a01 <> 0 Then
    If a05 = a01 Then
        Label1(1).Caption = 2 * a05
        S = S + 2 * a05
        Label1(5).Caption = 0
    End If
End If

If a13 <> 0 And a09 = 0 And a05 <> 0 And a01 = 0 Then
    If a13 = a05 Then
        Label1(1).Caption = 2 * a13
        S = S + 2 * a13
        Label1(13).Caption = 0
        Label1(5).Caption = 0
    Else
        Label1(1).Caption = a05
        Label1(5).Caption = a13
        Label1(13).Caption = 0
    End If
End If

If a13 = 0 And a09 <> 0 And a05 = 0 And a01 <> 0 Then
    If a09 = a01 Then
        Label1(1).Caption = 2 * a09
        S = S + 2 * a09
        Label1(9).Caption = 0
    Else
        Label1(5).Caption = a09
        Label1(9).Caption = 0
    End If
End If

If a13 <> 0 And a09 = 0 And a05 = 0 And a01 <> 0 Then
    If a13 = a01 Then
        Label1(1).Caption = 2 * a13
        S = S + 2 * a13
        Label1(13).Caption = 0
    Else
        Label1(5).Caption = a13
        Label1(13).Caption = 0
    End If
End If
'----华丽的分割线----
If a13 <> 0 And a09 <> 0 And a05 <> 0 And a01 = 0 Then
    If a13 = a09 And a09 <> a05 Then
        Label1(1).Caption = a05
        Label1(5).Caption = 2 * a13
        S = S + 2 * a13
        Label1(9).Caption = 0
        Label1(13).Caption = 0
    Else
        If a09 = a05 And a13 <> a09 Then
            Label1(1).Caption = 2 * a05
            S = S + 2 * a05
            Label1(5).Caption = a13
            Label1(9).Caption = 0
            Label1(13).Caption = 0
        Else
            If a13 = a09 And a13 = a05 Then
                Label1(1).Caption = 2 * a05
                S = S + 2 * a05
                Label1(5).Caption = a13
                Label1(9).Caption = 0
                Label1(13).Caption = 0
            Else
                Label1(1).Caption = a05
                Label1(5).Caption = a09
                Label1(9).Caption = a13
                Label1(13).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a13 = 0 And a09 <> 0 And a05 <> 0 And a01 <> 0 Then
    If a09 = a05 And a05 <> a01 Then
        Label1(5).Caption = 2 * a09
        S = S + 2 * a09
        Label1(9).Caption = 0
        Label1(13).Caption = 0
    Else
        If a05 = a01 And a09 <> a05 Then
            Label1(1).Caption = 2 * a05
            S = S + 2 * a05
            Label1(5).Caption = a09
            Label1(9).Caption = 0
            Label1(13).Caption = 0
        Else
            If a09 = a05 And a09 = a01 Then
                Label1(1).Caption = 2 * a05
                S = S + 2 * a05
                Label1(5).Caption = a09
                Label1(9).Caption = 0
                Label1(13).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a13 <> 0 And a09 <> 0 And a05 = 0 And a01 <> 0 Then
    If a13 = a09 And a13 <> a01 Then
        Label1(5).Caption = 2 * a09
        S = S + 2 * a09
        Label1(9).Caption = 0
        Label1(13).Caption = 0
    Else
        If a13 <> a09 And a09 = a01 Then
            Label1(1).Caption = 2 * a09
            S = S + 2 * a09
            Label1(5).Caption = a13
            Label1(9).Caption = 0
            Label1(13).Caption = 0
        Else
            If a13 = a09 And a09 = a01 Then
                Label1(1).Caption = 2 * a09
                S = S + 2 * a09
                Label1(5).Caption = a13
                Label1(9).Caption = 0
                Label1(13).Caption = 0
            Else
                Label1(5).Caption = a09
                Label1(9).Caption = a13
                Label1(13).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a13 <> 0 And a09 = 0 And a05 <> 0 And a01 <> 0 Then
    If a13 = a05 And a13 <> a01 Then
        Label1(5).Caption = 2 * a13
        S = S + 2 * a13
        Label1(9).Caption = 0
        Label1(13).Caption = 0
    Else
        If a13 <> a05 And a05 = a01 Then
            Label1(1).Caption = 2 * a05
            S = S + 2 * a05
            Label1(5).Caption = a13
            Label1(9).Caption = 0
            Label1(13).Caption = 0
        Else
            If a13 = a05 And a05 = a01 Then
                Label1(1).Caption = 2 * a05
                S = S + 2 * a05
                Label1(5).Caption = a13
                Label1(9).Caption = 0
                Label1(13).Caption = 0
            Else
                Label1(9).Caption = a13
                Label1(13).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a13 <> 0 And a09 <> 0 And a05 <> 0 And a01 <> 0 Then
    If a13 = a09 And a05 = a01 Then
         Label1(1).Caption = 2 * a05
         S = S + 2 * a05
         Label1(5).Caption = 2 * a13
         S = S + 2 * a13
         Label1(9).Caption = 0
         Label1(13).Caption = 0
    Else
        If a13 <> a09 And a09 = a05 And a05 <> a01 Then
            Label1(5).Caption = 2 * a09
            S = S + 2 * a09
            Label1(9).Caption = a13
            Label1(13).Caption = 0
        Else
            If a13 = a09 And a09 <> a05 Then
                Label1(9).Caption = 2 * a13
                S = S + 2 * a13
                Label1(13).Caption = 0
            Else
                If a09 <> a05 And a05 = a01 Then
                    Label1(1).Caption = 2 * a05
                    S = S + 2 * a05
                    Label1(5).Caption = a09
                    Label1(9).Caption = a13
                    Label1(13).Caption = 0
                Else
                    If a13 = a09 And a09 = a05 And a05 <> a01 Then
                        Label1(5).Caption = 2 * a09
                        S = S + 2 * a09
                        Label1(9).Caption = a13
                        Label1(13).Caption = 0
                    Else
                        If a13 <> a09 And a09 = a05 And a05 = a01 Then
                            Label1(1).Caption = 2 * a05
                            S = S + 2 * a05
                            Label1(5).Caption = a09
                            Label1(9).Caption = a13
                            Label1(13).Caption = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
'--------超级华丽的分割线--------
'--------超级华丽的分割线--------
If a14 <> 0 And a10 = 0 And a06 = 0 And a02 = 0 Then
    Label1(2).Caption = a14
    Label1(14).Caption = 0
End If

If a14 = 0 And a10 <> 0 And a06 = 0 And a02 = 0 Then
    Label1(2).Caption = a10
    Label1(10).Caption = 0
End If

If a14 = 0 And a10 = 0 And a06 <> 0 And a02 = 0 Then
    Label1(2).Caption = a06
    Label1(6).Caption = 0
End If

If a14 = 0 And a10 = 0 And a06 = 0 And a02 <> 0 Then

End If

'----华丽的分割线----

If a14 <> 0 And a10 <> 0 And a06 = 0 And a02 = 0 Then
    If a14 = a10 Then
        Label1(2).Caption = 2 * a14
        S = S + 2 * a14
        Label1(14).Caption = 0
        Label1(10).Caption = 0
    Else
        Label1(6).Caption = a14
        Label1(2).Caption = a10
        Label1(14).Caption = 0
        Label1(10).Caption = 0
    End If
End If

If a14 = 0 And a10 <> 0 And a06 <> 0 And a02 = 0 Then
    If a10 = a06 Then
        Label1(2).Caption = 2 * a10
        S = S + 2 * a10
        Label1(10).Caption = 0
        Label1(6).Caption = 0
    Else
        Label1(2).Caption = a06
        Label1(6).Caption = a10
        Label1(10).Caption = 0
    End If
End If

If a14 = 0 And a10 = 0 And a06 <> 0 And a02 <> 0 Then
    If a06 = a02 Then
        Label1(2).Caption = 2 * a06
        S = S + 2 * a06
        Label1(6).Caption = 0
    End If
End If

If a14 <> 0 And a10 = 0 And a06 <> 0 And a02 = 0 Then
    If a14 = a06 Then
        Label1(2).Caption = 2 * a14
        S = S + 2 * a14
        Label1(14).Caption = 0
        Label1(6).Caption = 0
    Else
        Label1(2).Caption = a06
        Label1(6).Caption = a14
        Label1(14).Caption = 0
    End If
End If

If a14 = 0 And a10 <> 0 And a06 = 0 And a02 <> 0 Then
    If a10 = a02 Then
        Label1(2).Caption = 2 * a10
        S = S + 2 * a10
        Label1(10).Caption = 0
    Else
        Label1(6).Caption = a10
        Label1(10).Caption = 0
    End If
End If

If a14 <> 0 And a10 = 0 And a06 = 0 And a02 <> 0 Then
    If a14 = a02 Then
        Label1(2).Caption = 2 * a14
        S = S + 2 * a14
        Label1(14).Caption = 0
    Else
        Label1(6).Caption = a14
        Label1(14).Caption = 0
    End If
End If
'----华丽的分割线----
If a14 <> 0 And a10 <> 0 And a06 <> 0 And a02 = 0 Then
    If a14 = a10 And a10 <> a06 Then
        Label1(2).Caption = a06
        Label1(6).Caption = 2 * a14
        S = S + 2 * a14
        Label1(10).Caption = 0
        Label1(14).Caption = 0
    Else
        If a10 = a06 And a14 <> a10 Then
            Label1(2).Caption = 2 * a06
            S = S + 2 * a06
            Label1(6).Caption = a14
            Label1(10).Caption = 0
            Label1(14).Caption = 0
        Else
            If a14 = a10 And a14 = a06 Then
                Label1(2).Caption = 2 * a06
                S = S + 2 * a06
                Label1(6).Caption = a14
                Label1(10).Caption = 0
                Label1(14).Caption = 0
            Else
                Label1(2).Caption = a06
                Label1(6).Caption = a10
                Label1(10).Caption = a14
                Label1(14).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a14 = 0 And a10 <> 0 And a06 <> 0 And a02 <> 0 Then
    If a10 = a06 And a06 <> a02 Then
        Label1(6).Caption = 2 * a10
        S = S + 2 * a10
        Label1(10).Caption = 0
        Label1(14).Caption = 0
    Else
        If a06 = a02 And a10 <> a06 Then
            Label1(2).Caption = 2 * a06
            S = S + 2 * a06
            Label1(6).Caption = a10
            Label1(10).Caption = 0
            Label1(14).Caption = 0
        Else
            If a10 = a06 And a10 = a02 Then
                Label1(2).Caption = 2 * a06
                S = S + 2 * a06
                Label1(6).Caption = a10
                Label1(10).Caption = 0
                Label1(14).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a14 <> 0 And a10 <> 0 And a06 = 0 And a02 <> 0 Then
    If a14 = a10 And a14 <> a02 Then
        Label1(6).Caption = 2 * a10
        S = S + 2 * a10
        Label1(10).Caption = 0
        Label1(14).Caption = 0
    Else
        If a14 <> a10 And a10 = a02 Then
            Label1(2).Caption = 2 * a10
            S = S + 2 * a10
            Label1(6).Caption = a14
            Label1(10).Caption = 0
            Label1(14).Caption = 0
        Else
            If a14 = a10 And a10 = a02 Then
                Label1(2).Caption = 2 * a10
                S = S + 2 * a10
                Label1(6).Caption = a14
                Label1(10).Caption = 0
                Label1(14).Caption = 0
            Else
                Label1(6).Caption = a10
                Label1(10).Caption = a14
                Label1(14).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a14 <> 0 And a10 = 0 And a06 <> 0 And a02 <> 0 Then
    If a14 = a06 And a14 <> a02 Then
        Label1(6).Caption = 2 * a14
        S = S + 2 * a14
        Label1(10).Caption = 0
        Label1(14).Caption = 0
    Else
        If a14 <> a06 And a06 = a02 Then
            Label1(2).Caption = 2 * a06
            S = S + 2 * a06
            Label1(6).Caption = a14
            Label1(10).Caption = 0
            Label1(14).Caption = 0
        Else
            If a14 = a06 And a06 = a02 Then
                Label1(2).Caption = 2 * a06
                S = S + 2 * a06
                Label1(6).Caption = a14
                Label1(10).Caption = 0
                Label1(14).Caption = 0
            Else
                Label1(10).Caption = a14
                Label1(14).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a14 <> 0 And a10 <> 0 And a06 <> 0 And a02 <> 0 Then
    If a14 = a10 And a06 = a02 Then
         Label1(2).Caption = 2 * a06
         S = S + 2 * a06
         Label1(6).Caption = 2 * a14
         S = S + 2 * a14
         Label1(10).Caption = 0
         Label1(14).Caption = 0
    Else
        If a14 <> a10 And a10 = a06 And a06 <> a02 Then
            Label1(6).Caption = 2 * a10
            S = S + 2 * a10
            Label1(10).Caption = a14
            Label1(14).Caption = 0
        Else
            If a14 = a10 And a10 <> a06 Then
                Label1(10).Caption = 2 * a14
                S = S + 2 * a14
                Label1(14).Caption = 0
            Else
                If a10 <> a06 And a06 = a02 Then
                    Label1(2).Caption = 2 * a06
                    S = S + 2 * a06
                    Label1(6).Caption = a10
                    Label1(10).Caption = a14
                    Label1(14).Caption = 0
                Else
                    If a14 = a10 And a10 = a06 And a06 <> a02 Then
                        Label1(6).Caption = 2 * a10
                        S = S + 2 * a10
                        Label1(10).Caption = a14
                        Label1(14).Caption = 0
                    Else
                        If a14 <> a10 And a10 = a06 And a06 = a02 Then
                            Label1(2).Caption = 2 * a06
                            S = S + 2 * a06
                            Label1(6).Caption = a10
                            Label1(10).Caption = a14
                            Label1(14).Caption = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
'--------超级华丽的分割线--------
'--------超级华丽的分割线--------
Call Command66_Click

If a15 <> 0 And a11 = 0 And a07 = 0 And a03 = 0 Then
    Label1(3).Caption = a15
    Label1(15).Caption = 0
End If

If a15 = 0 And a11 <> 0 And a07 = 0 And a03 = 0 Then
    Label1(3).Caption = a11
    Label1(11).Caption = 0
End If

If a15 = 0 And a11 = 0 And a07 <> 0 And a03 = 0 Then
    Label1(3).Caption = a07
    Label1(7).Caption = 0
End If

If a15 = 0 And a11 = 0 And a07 = 0 And a03 <> 0 Then

End If

'----华丽的分割线----

If a15 <> 0 And a11 <> 0 And a07 = 0 And a03 = 0 Then
    If a15 = a11 Then
        Label1(3).Caption = 2 * a15
        S = S + 2 * a15
        Label1(15).Caption = 0
        Label1(11).Caption = 0
    Else
        Label1(7).Caption = a15
        Label1(3).Caption = a11
        Label1(15).Caption = 0
        Label1(11).Caption = 0
    End If
End If

If a15 = 0 And a11 <> 0 And a07 <> 0 And a03 = 0 Then
    If a11 = a07 Then
        Label1(3).Caption = 2 * a11
        S = S + 2 * a11
        Label1(11).Caption = 0
        Label1(7).Caption = 0
    Else
        Label1(3).Caption = a07
        Label1(7).Caption = a11
        Label1(11).Caption = 0
    End If
End If

If a15 = 0 And a11 = 0 And a07 <> 0 And a03 <> 0 Then
    If a07 = a03 Then
        Label1(3).Caption = 2 * a07
        S = S + 2 * a07
        Label1(7).Caption = 0
    End If
End If

If a15 <> 0 And a11 = 0 And a07 <> 0 And a03 = 0 Then
    If a15 = a07 Then
        Label1(3).Caption = 2 * a15
        S = S + 2 * a15
        Label1(15).Caption = 0
        Label1(7).Caption = 0
    Else
        Label1(3).Caption = a07
        Label1(7).Caption = a15
        Label1(15).Caption = 0
    End If
End If

If a15 = 0 And a11 <> 0 And a07 = 0 And a03 <> 0 Then
    If a11 = a03 Then
        Label1(3).Caption = 2 * a11
        S = S + 2 * a11
        Label1(11).Caption = 0
    Else
        Label1(7).Caption = a11
        Label1(11).Caption = 0
    End If
End If

If a15 <> 0 And a11 = 0 And a07 = 0 And a03 <> 0 Then
    If a15 = a03 Then
        Label1(3).Caption = 2 * a15
        S = S + 2 * a15
        Label1(15).Caption = 0
    Else
        Label1(7).Caption = a15
        Label1(15).Caption = 0
    End If
End If
'----华丽的分割线----
If a15 <> 0 And a11 <> 0 And a07 <> 0 And a03 = 0 Then
    If a15 = a11 And a11 <> a07 Then
        Label1(3).Caption = a07
        Label1(7).Caption = 2 * a15
        S = S + 2 * a15
        Label1(11).Caption = 0
        Label1(15).Caption = 0
    Else
        If a11 = a07 And a15 <> a11 Then
            Label1(3).Caption = 2 * a07
            S = S + 2 * a07
            Label1(7).Caption = a15
            Label1(11).Caption = 0
            Label1(15).Caption = 0
        Else
            If a15 = a11 And a15 = a07 Then
                Label1(3).Caption = 2 * a07
                S = S + 2 * a07
                Label1(7).Caption = a15
                Label1(11).Caption = 0
                Label1(15).Caption = 0
            Else
                Label1(3).Caption = a07
                Label1(7).Caption = a11
                Label1(11).Caption = a15
                Label1(15).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a15 = 0 And a11 <> 0 And a07 <> 0 And a03 <> 0 Then
    If a11 = a07 And a07 <> a03 Then
        Label1(7).Caption = 2 * a11
        S = S + 2 * a11
        Label1(11).Caption = 0
        Label1(15).Caption = 0
    Else
        If a07 = a03 And a11 <> a07 Then
            Label1(3).Caption = 2 * a07
            S = S + 2 * a07
            Label1(7).Caption = a11
            Label1(11).Caption = 0
            Label1(15).Caption = 0
        Else
            If a11 = a07 And a11 = a03 Then
                Label1(3).Caption = 2 * a07
                S = S + 2 * a07
                Label1(7).Caption = a11
                Label1(11).Caption = 0
                Label1(15).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a15 <> 0 And a11 <> 0 And a07 = 0 And a03 <> 0 Then
    If a15 = a11 And a15 <> a03 Then
        Label1(7).Caption = 2 * a11
        S = S + 2 * a11
        Label1(11).Caption = 0
        Label1(15).Caption = 0
    Else
        If a15 <> a11 And a11 = a03 Then
            Label1(3).Caption = 2 * a11
            S = S + 2 * a11
            Label1(7).Caption = a15
            Label1(11).Caption = 0
            Label1(15).Caption = 0
        Else
            If a15 = a11 And a11 = a03 Then
                Label1(3).Caption = 2 * a11
                S = S + 2 * a11
                Label1(7).Caption = a15
                Label1(11).Caption = 0
                Label1(15).Caption = 0
            Else
                Label1(7).Caption = a11
                Label1(11).Caption = a15
                Label1(15).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a15 <> 0 And a11 = 0 And a07 <> 0 And a03 <> 0 Then
    If a15 = a07 And a15 <> a03 Then
        Label1(7).Caption = 2 * a15
        S = S + 2 * a15
        Label1(11).Caption = 0
        Label1(15).Caption = 0
    Else
        If a15 <> a07 And a07 = a03 Then
            Label1(3).Caption = 2 * a07
            S = S + 2 * a07
            Label1(7).Caption = a15
            Label1(11).Caption = 0
            Label1(15).Caption = 0
        Else
            If a15 = a07 And a07 = a03 Then
                Label1(3).Caption = 2 * a07
                S = S + 2 * a07
                Label1(7).Caption = a15
                Label1(11).Caption = 0
                Label1(15).Caption = 0
            Else
                Label1(11).Caption = a15
                Label1(15).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a15 <> 0 And a11 <> 0 And a07 <> 0 And a03 <> 0 Then
    If a15 = a11 And a07 = a03 Then
         Label1(3).Caption = 2 * a07
         S = S + 2 * a07
         Label1(7).Caption = 2 * a15
         S = S + 2 * a15
         Label1(11).Caption = 0
         Label1(15).Caption = 0
    Else
        If a15 <> a11 And a11 = a07 And a07 <> a03 Then
            Label1(7).Caption = 2 * a11
            S = S + 2 * a11
            Label1(11).Caption = a15
            Label1(15).Caption = 0
        Else
            If a15 = a11 And a11 <> a07 Then
                Label1(11).Caption = 2 * a15
                S = S + 2 * a15
                Label1(15).Caption = 0
            Else
                If a11 <> a07 And a07 = a03 Then
                    Label1(3).Caption = 2 * a07
                    S = S + 2 * a07
                    Label1(7).Caption = a11
                    Label1(11).Caption = a15
                    Label1(15).Caption = 0
                Else
                    If a15 = a11 And a11 = a07 And a07 <> a03 Then
                        Label1(7).Caption = 2 * a11
                        S = S + 2 * a11
                        Label1(11).Caption = a15
                        Label1(15).Caption = 0
                    Else
                        If a15 <> a11 And a11 = a07 And a07 = a03 Then
                            Label1(3).Caption = 2 * a07
                            S = S + 2 * a07
                            Label1(7).Caption = a11
                            Label1(11).Caption = a15
                            Label1(15).Caption = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
'--------超级华丽的分割线--------
'--------超级华丽的分割线--------
If a16 <> 0 And a12 = 0 And a08 = 0 And a04 = 0 Then
    Label1(4).Caption = a16
    Label1(16).Caption = 0
End If

If a16 = 0 And a12 <> 0 And a08 = 0 And a04 = 0 Then
    Label1(4).Caption = a12
    Label1(12).Caption = 0
End If

If a16 = 0 And a12 = 0 And a08 <> 0 And a04 = 0 Then
    Label1(4).Caption = a08
    Label1(8).Caption = 0
End If

If a16 = 0 And a12 = 0 And a08 = 0 And a04 <> 0 Then

End If

'----华丽的分割线----

If a16 <> 0 And a12 <> 0 And a08 = 0 And a04 = 0 Then
    If a16 = a12 Then
        Label1(4).Caption = 2 * a16
        S = S + 2 * a16
        Label1(16).Caption = 0
        Label1(12).Caption = 0
    Else
        Label1(8).Caption = a16
        Label1(4).Caption = a12
        Label1(16).Caption = 0
        Label1(12).Caption = 0
    End If
End If

If a16 = 0 And a12 <> 0 And a08 <> 0 And a04 = 0 Then
    If a12 = a08 Then
        Label1(4).Caption = 2 * a12
        S = S + 2 * a12
        Label1(12).Caption = 0
        Label1(8).Caption = 0
    Else
        Label1(4).Caption = a08
        Label1(8).Caption = a12
        Label1(12).Caption = 0
    End If
End If

If a16 = 0 And a12 = 0 And a08 <> 0 And a04 <> 0 Then
    If a08 = a04 Then
        Label1(4).Caption = 2 * a08
        S = S + 2 * a08
        Label1(8).Caption = 0
    End If
End If

If a16 <> 0 And a12 = 0 And a08 <> 0 And a04 = 0 Then
    If a16 = a08 Then
        Label1(4).Caption = 2 * a16
        S = S + 2 * a16
        Label1(16).Caption = 0
        Label1(8).Caption = 0
    Else
        Label1(4).Caption = a08
        Label1(8).Caption = a16
        Label1(16).Caption = 0
    End If
End If

If a16 = 0 And a12 <> 0 And a08 = 0 And a04 <> 0 Then
    If a12 = a04 Then
        Label1(4).Caption = 2 * a12
        S = S + 2 * a12
        Label1(12).Caption = 0
    Else
        Label1(8).Caption = a12
        Label1(12).Caption = 0
    End If
End If

If a16 <> 0 And a12 = 0 And a08 = 0 And a04 <> 0 Then
    If a16 = a04 Then
        Label1(4).Caption = 2 * a16
        S = S + 2 * a16
        Label1(16).Caption = 0
    Else
        Label1(8).Caption = a16
        Label1(16).Caption = 0
    End If
End If
'----华丽的分割线----
If a16 <> 0 And a12 <> 0 And a08 <> 0 And a04 = 0 Then
    If a16 = a12 And a12 <> a08 Then
        Label1(4).Caption = a08
        Label1(8).Caption = 2 * a16
        S = S + 2 * a16
        Label1(12).Caption = 0
        Label1(16).Caption = 0
    Else
        If a12 = a08 And a16 <> a12 Then
            Label1(4).Caption = 2 * a08
            S = S + 2 * a08
            Label1(8).Caption = a16
            Label1(12).Caption = 0
            Label1(16).Caption = 0
        Else
            If a16 = a12 And a16 = a08 Then
                Label1(4).Caption = 2 * a08
                S = S + 2 * a08
                Label1(8).Caption = a16
                Label1(12).Caption = 0
                Label1(16).Caption = 0
            Else
                Label1(4).Caption = a08
                Label1(8).Caption = a12
                Label1(12).Caption = a16
                Label1(16).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a16 = 0 And a12 <> 0 And a08 <> 0 And a04 <> 0 Then
    If a12 = a08 And a08 <> a04 Then
        Label1(8).Caption = 2 * a12
        S = S + 2 * a12
        Label1(12).Caption = 0
        Label1(16).Caption = 0
    Else
        If a08 = a04 And a12 <> a08 Then
            Label1(4).Caption = 2 * a08
            S = S + 2 * a08
            Label1(8).Caption = a12
            Label1(12).Caption = 0
            Label1(16).Caption = 0
        Else
            If a12 = a08 And a12 = a04 Then
                Label1(4).Caption = 2 * a08
                S = S + 2 * a08
                Label1(8).Caption = a12
                Label1(12).Caption = 0
                Label1(16).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a16 <> 0 And a12 <> 0 And a08 = 0 And a04 <> 0 Then
    If a16 = a12 And a16 <> a04 Then
        Label1(8).Caption = 2 * a12
        S = S + 2 * a12
        Label1(12).Caption = 0
        Label1(16).Caption = 0
    Else
        If a16 <> a12 And a12 = a04 Then
            Label1(4).Caption = 2 * a12
            S = S + 2 * a12
            Label1(8).Caption = a16
            Label1(12).Caption = 0
            Label1(16).Caption = 0
        Else
            If a16 = a12 And a12 = a04 Then
                Label1(4).Caption = 2 * a12
                S = S + 2 * a12
                Label1(8).Caption = a16
                Label1(12).Caption = 0
                Label1(16).Caption = 0
            Else
                Label1(8).Caption = a12
                Label1(12).Caption = a16
                Label1(16).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a16 <> 0 And a12 = 0 And a08 <> 0 And a04 <> 0 Then
    If a16 = a08 And a16 <> a04 Then
        Label1(8).Caption = 2 * a16
        S = S + 2 * a16
        Label1(12).Caption = 0
        Label1(16).Caption = 0
    Else
        If a16 <> a08 And a08 = a04 Then
            Label1(4).Caption = 2 * a08
            S = S + 2 * a08
            Label1(8).Caption = a16
            Label1(12).Caption = 0
            Label1(16).Caption = 0
        Else
            If a16 = a08 And a08 = a04 Then
                Label1(4).Caption = 2 * a08
                S = S + 2 * a08
                Label1(8).Caption = a16
                Label1(12).Caption = 0
                Label1(16).Caption = 0
            Else
                Label1(12).Caption = a16
                Label1(16).Caption = 0
            End If
        End If
    End If
End If
'----华丽的分割线----
If a16 <> 0 And a12 <> 0 And a08 <> 0 And a04 <> 0 Then
    If a16 = a12 And a08 = a04 Then
         Label1(4).Caption = 2 * a08
         S = S + 2 * a08
         Label1(8).Caption = 2 * a16
         S = S + 2 * a16
         Label1(12).Caption = 0
         Label1(16).Caption = 0
    Else
        If a16 <> a12 And a12 = a08 And a08 <> a04 Then
            Label1(8).Caption = 2 * a12
            S = S + 2 * a12
            Label1(12).Caption = a16
            Label1(16).Caption = 0
        Else
            If a16 = a12 And a12 <> a08 Then
                Label1(12).Caption = 2 * a16
                S = S + 2 * a16
                Label1(16).Caption = 0
            Else
                If a12 <> a08 And a08 = a04 Then
                    Label1(4).Caption = 2 * a08
                    S = S + 2 * a08
                    Label1(8).Caption = a12
                    Label1(12).Caption = a16
                    Label1(16).Caption = 0
                Else
                    If a16 = a12 And a12 = a08 And a08 <> a04 Then
                        Label1(8).Caption = 2 * a12
                        S = S + 2 * a12
                        Label1(12).Caption = a16
                        Label1(16).Caption = 0
                    Else
                        If a16 <> a12 And a12 = a08 And a08 = a04 Then
                            Label1(4).Caption = 2 * a08
                            S = S + 2 * a08
                            Label1(8).Caption = a12
                            Label1(12).Caption = a16
                            Label1(16).Caption = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
Call Command101_Click
Call Command5_Click
Call Command6_Click
End Sub


Private Sub Command5_Click()
Randomize
X = Val(Int(Rnd * 11)) + Val(1)
If X = 11 Then
    Y = 4
Else
    Y = 2
End If
Do
Randomize
Z = Val(Int(Rnd * 16)) + Val(1)
Loop Until Label1(Z).Caption = 0
Label1(Z).Caption = Y

End Sub

Private Sub Command6_Click()
i = 1
Do
If Label1(i) = 0 Then
    Me.Image1(i).Picture = LoadPicture(App.Path + "\0.gif")
Else
    If Label1(i) = 2 Then
        Me.Image1(i).Picture = LoadPicture(App.Path + "\2.gif")
    Else
        If Label1(i) = 4 Then
            Me.Image1(i).Picture = LoadPicture(App.Path + "\4.gif")
        Else
            If Label1(i) = 8 Then
                Me.Image1(i).Picture = LoadPicture(App.Path + "\8.gif")
            Else
                If Label1(i) = 16 Then
                    Me.Image1(i).Picture = LoadPicture(App.Path + "\16.gif")
                Else
                    If Label1(i) = 32 Then
                        Me.Image1(i).Picture = LoadPicture(App.Path + "\32.gif")
                    Else
                        If Label1(i) = 64 Then
                            Me.Image1(i).Picture = LoadPicture(App.Path + "\64.gif")
                        Else
                            If Label1(i) = 128 Then
                                Me.Image1(i).Picture = LoadPicture(App.Path + "\128.gif")
                            Else
                                If Label1(i) = 256 Then
                                    Me.Image1(i).Picture = LoadPicture(App.Path + "\256.gif")
                                Else
                                    If Label1(i) = 512 Then
                                        Me.Image1(i).Picture = LoadPicture(App.Path + "\512.gif")
                                    Else
                                        If Label1(i) = 1024 Then
                                            Me.Image1(i).Picture = LoadPicture(App.Path + "\1024.gif")
                                        Else
                                            If Label1(i) = 2048 Then
                                                Me.Image1(i).Picture = LoadPicture(App.Path + "\2048.gif")
                                            Else
                                                If Label1(i) = 4096 Then
                                                    Me.Image1(i).Picture = LoadPicture(App.Path + "\4096.gif")
                                                Else
                                                    If Label1(i) = 8192 Then
                                                        Me.Image1(i).Picture = LoadPicture(App.Path + "\8192.gif")
                                                    Else
                                                        If Label1(i) = 16384 Then
                                                            Me.Image1(i).Picture = LoadPicture(App.Path + "\16384.gif")
                                                        Else
                                                            If Label1(i) = 32768 Then
                                                                Me.Image1(i).Picture = LoadPicture(App.Path + "\32768.gif")
                                                            Else
                                                                If Label1(i) = 65536 Then
                                                                    Me.Image1(i).Picture = LoadPicture(App.Path + "\65536.gif")
                                                                Else
                                                                    If Label1(i) = 131072 Then
                                                                        Me.Image1(i).Picture = LoadPicture(App.Path + "\131072.gif")
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
i = i + 1
Loop Until i = 17
        
End Sub

Private Sub Command66_Click() '化简
a01 = Label1(1).Caption
a02 = Label1(2).Caption
a03 = Label1(3).Caption
a04 = Label1(4).Caption
a05 = Label1(5).Caption
a06 = Label1(6).Caption
a07 = Label1(7).Caption
a08 = Label1(8).Caption
a09 = Label1(9).Caption
a10 = Label1(10).Caption
a11 = Label1(11).Caption
a12 = Label1(12).Caption
a13 = Label1(13).Caption
a14 = Label1(14).Caption
a15 = Label1(15).Caption
a16 = Label1(16).Caption

End Sub

Private Sub Form_Load()
Command5_Click
Command5_Click

Call Command6_Click
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
    Call Command4_Click
End If
If KeyCode = 40 Then
    Call Command2_Click
End If
If KeyCode = 37 Then
    Call Command3_Click
End If
If KeyCode = 39 Then
    Call Command1_Click
End If

End Sub


