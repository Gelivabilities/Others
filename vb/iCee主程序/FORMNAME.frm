VERSION 5.00
Begin VB.Form FORMNAME 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00231C09&
   BorderStyle     =   0  'None
   Caption         =   "������"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8565
   ForeColor       =   &H00000000&
   Icon            =   "FORMNAME.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   571
   StartUpPosition =   3  '����ȱʡ
   Begin ICEE.ICHECK cHECK1 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   58
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
   End
   Begin ICEE.ICEE_KEY cmdGoPath 
      Height          =   495
      Left            =   7200
      TabIndex        =   57
      Top             =   8640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
   End
   Begin ICEE.ICEE_KEY cmdSelect 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   53
      Top             =   7320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
   End
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7800
      Picture         =   "FORMNAME.frx":038A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   8
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7800
      Picture         =   "FORMNAME.frx":046E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   7
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7800
      Picture         =   "FORMNAME.frx":0552
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   6
      Top             =   15
      Width           =   750
   End
   Begin VB.TextBox txtPath 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   320
      Left            =   240
      TabIndex        =   5
      Top             =   8760
      Width           =   6855
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   4890
      Left            =   240
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   2400
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      ItemData        =   "FORMNAME.frx":0636
      Left            =   240
      List            =   "FORMNAME.frx":065B
      TabIndex        =   2
      Text            =   "*.*"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.PictureBox PD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00231C09&
      BorderStyle     =   0  'None
      Height          =   7695
      Left            =   3375
      ScaleHeight     =   513
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   337
      TabIndex        =   0
      Top             =   840
      Width           =   5055
      Begin ICEE.IVScroll SCRO 
         Height          =   7695
         Left            =   4725
         TabIndex        =   9
         Top             =   0
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   13573
         MinV            =   0
         MaxV            =   20
         Value           =   0
         SmallChange     =   1
         LargeChange     =   10
      End
      Begin VB.PictureBox PO 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00231C09&
         BorderStyle     =   0  'None
         Height          =   9855
         Left            =   0
         ScaleHeight     =   657
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   305
         TabIndex        =   1
         Top             =   0
         Width           =   4575
         Begin VB.TextBox txtStep 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   3360
            TabIndex        =   52
            Text            =   "1"
            Top             =   660
            Width           =   855
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Text            =   "�Զ���"
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox txtRemoveStr 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   320
            Left            =   360
            TabIndex        =   48
            Text            =   "0000"
            Top             =   6840
            Width           =   2415
         End
         Begin ICEE.ICEE_KEY cmdChangeExt 
            Height          =   495
            Left            =   3120
            TabIndex        =   47
            Top             =   8160
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
         End
         Begin VB.TextBox txtNewExt 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   360
            TabIndex        =   46
            Text            =   "0011"
            ToolTipText     =   "ֻҪ������չ��������Ҫǰ��ġ�.������"
            Top             =   8280
            Width           =   2415
         End
         Begin VB.PictureBox PF 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00835C02&
            BorderStyle     =   0  'None
            Height          =   1335
            Index           =   1
            Left            =   3000
            ScaleHeight     =   89
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   89
            TabIndex        =   38
            Top             =   4080
            Width           =   1335
            Begin VB.OptionButton Option2 
               BackColor       =   &H00231C09&
               Caption         =   "ɾ��"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   40
               Top             =   120
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.OptionButton Option2 
               BackColor       =   &H00231C09&
               Caption         =   "ת��Ϊ_"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   39
               Top             =   480
               Width           =   975
            End
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00231C09&
            Caption         =   "ǰ׺"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   36
            Top             =   4200
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00231C09&
            Caption         =   "��׺"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   35
            Top             =   4560
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.PictureBox PF 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00835C02&
            BorderStyle     =   0  'None
            Height          =   1095
            Index           =   0
            Left            =   1200
            ScaleHeight     =   73
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   113
            TabIndex        =   30
            Top             =   4080
            Width           =   1695
            Begin VB.OptionButton Option1 
               BackColor       =   &H00231C09&
               Caption         =   "��д"
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   33
               Top             =   120
               Width           =   735
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00231C09&
               Caption         =   "Сд"
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   32
               Top             =   480
               Width           =   735
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00231C09&
               Caption         =   "����ĸ��д"
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   2
               Left            =   120
               TabIndex        =   31
               Top             =   840
               Value           =   -1  'True
               Width           =   1200
            End
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00231C09&
            Caption         =   "�滻ǰ׺"
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   28
            Top             =   1440
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00231C09&
            Caption         =   "�滻��׺"
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   27
            Top             =   1800
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00231C09&
            Caption         =   "��ǰ׺��ǰ��"
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   26
            Top             =   1440
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00231C09&
            Caption         =   "��ǰ׺�ĺ���"
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   3
            Left            =   1320
            TabIndex        =   25
            Top             =   1800
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00231C09&
            Caption         =   "�ں�׺��ǰ��"
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   4
            Left            =   2760
            TabIndex        =   24
            Top             =   1440
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00231C09&
            Caption         =   "�ں�׺�ĺ���"
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   5
            Left            =   2760
            TabIndex        =   23
            Top             =   1800
            Width           =   1455
         End
         Begin ICEE.ICEE_KEY COMMAND2 
            Height          =   495
            Left            =   3120
            TabIndex        =   20
            Top             =   3000
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY COMMAND1 
            Height          =   495
            Left            =   3120
            TabIndex        =   21
            Top             =   5640
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY cmdRemoveStr 
            Height          =   495
            Left            =   3120
            TabIndex        =   49
            Top             =   6720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
         End
         Begin VB.Shape SB 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   495
            Index           =   3
            Left            =   120
            Top             =   8160
            Width           =   3015
         End
         Begin VB.Shape SB 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   495
            Index           =   1
            Left            =   120
            Top             =   6720
            Width           =   3015
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ȥ�����ļ������ַ�"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   12
            Left            =   120
            TabIndex        =   50
            Top             =   6240
            Width           =   1620
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   11
            Left            =   120
            TabIndex        =   45
            Top             =   7080
            Width           =   90
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Զ����ı�"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   0
            TabIndex        =   44
            Top             =   120
            Width           =   1050
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ļ��ַ�����"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   10
            Left            =   0
            TabIndex        =   43
            Top             =   3120
            Width           =   1260
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�޸���չ��"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   9
            Left            =   0
            TabIndex        =   42
            Top             =   7560
            Width           =   1050
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ո���"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   8
            Left            =   2880
            TabIndex        =   41
            Top             =   3840
            Width           =   720
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Χ"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   7
            Left            =   0
            TabIndex        =   37
            Top             =   3840
            Width           =   360
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ת����"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   6
            Left            =   1200
            TabIndex        =   34
            Top             =   3840
            Width           =   540
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Զ����ı���λ��"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   5
            Left            =   120
            TabIndex        =   29
            Top             =   1080
            Width           =   1440
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   2760
            TabIndex        =   22
            Top             =   660
            Width           =   450
         End
         Begin VB.Label lblSource 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "model.gif"
            ForeColor       =   &H00E0E0E0&
            Height          =   180
            Index           =   1
            Left            =   1320
            TabIndex        =   19
            Top             =   2640
            Width           =   810
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ת�����Ϊ��"
            ForeColor       =   &H00E0E0E0&
            Height          =   180
            Index           =   4
            Left            =   2520
            TabIndex        =   18
            Top             =   2400
            Width           =   1080
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ԭ�ļ���Ϊ��"
            ForeColor       =   &H00E0E0E0&
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   2400
            Width           =   1080
         End
         Begin VB.Label lblSource 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "example.gif"
            ForeColor       =   &H00E0E0E0&
            Height          =   180
            Index           =   0
            Left            =   1320
            TabIndex        =   16
            Top             =   2400
            Width           =   990
         End
         Begin VB.Label lblTarget 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1.gif"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   0
            Left            =   3600
            TabIndex        =   15
            Top             =   2400
            Width           =   450
         End
         Begin VB.Label lblTarget 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2.gif"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   1
            Left            =   3600
            TabIndex        =   14
            Top             =   2640
            Width           =   450
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ԭ�ļ���Ϊ��"
            ForeColor       =   &H00E0E0E0&
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   13
            Top             =   5640
            Width           =   1080
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ת�����Ϊ��"
            ForeColor       =   &H00E0E0E0&
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   12
            Top             =   5880
            Width           =   1080
         End
         Begin VB.Label lblS1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exa mpL e.exE"
            ForeColor       =   &H00E0E0E0&
            Height          =   180
            Left            =   1200
            TabIndex        =   11
            Top             =   5640
            Width           =   1170
         End
         Begin VB.Label lblT1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Example.exe"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   1200
            TabIndex        =   10
            Top             =   5880
            Width           =   990
         End
         Begin VB.Shape SB 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   495
            Index           =   2
            Left            =   0
            Top             =   480
            Width           =   4455
         End
      End
   End
   Begin ICEE.ICEE_KEY cmdSelect 
      Height          =   495
      Index           =   1
      Left            =   2040
      TabIndex        =   54
      Top             =   7320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
   End
   Begin ICEE.ICEE_KEY cmdSelect 
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   55
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
   End
   Begin ICEE.ICEE_KEY cmdSelect 
      Height          =   495
      Index           =   3
      Left            =   2040
      TabIndex        =   56
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
   End
   Begin ICEE.ICHECK cHECK1 
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   59
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
   End
   Begin ICEE.ICHECK cHECK1 
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   60
      Top             =   2040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
   End
   Begin ICEE.ICHECK cHECK1 
      Height          =   375
      Index           =   3
      Left            =   1560
      TabIndex        =   61
      Top             =   2040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
   End
   Begin VB.Shape SB 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   0
      Left            =   120
      Top             =   8640
      Width           =   7095
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ļ���չ������"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1260
   End
End
Attribute VB_Name = "FORMNAME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'

Option Explicit

Private Const LB_SETSEL = &H185
Private mCount As Integer           '������������

Public Sub ShowListItemToolTipText(ListX As VB.FileListBox, iY As Single)
    Dim YPos As Integer
    Dim FormX As VB.Form
    Dim FontX As New stdole.StdFont       '��������
      
    Set FormX = ListX.Container
    
    FontX.Bold = Me.Font.Bold
    FontX.Charset = Me.Font.Charset
    FontX.Italic = Me.Font.Italic
    FontX.name = Me.Font.name
    FontX.Size = Me.Font.Size
    FontX.Strikethrough = Me.Font.Strikethrough
    FontX.Underline = Me.Font.Underline
    FontX.Weight = Me.Font.Weight
      
    Set FormX.Font = ListX.Font
    YPos = iY \ FormX.TextHeight("Xyz") + ListX.TopIndex
    Set FormX.Font = FontX         '�ָ�����
      
    If YPos < ListX.ListCount And FormX.TextWidth(ListX.List(YPos)) > ListX.Width - 450 Then
        ListX.ToolTipText = ListX.List(YPos)
    Else
        ListX.ToolTipText = ""
    End If
End Sub

Private Sub cHECK1_Click(Index As Integer)
  Select Case Index
  Case 0
    File1.Normal = cHECK1(0).Value
  Case 1
    File1.ReadOnly = cHECK1(1).Value
  Case 2
    File1.Hidden = cHECK1(2).Value
  Case 3
    File1.System = cHECK1(3).Value
  End Select
End Sub

Private Sub Check2_Click(Index As Integer)
  '�ı� ת����Χ�����´���
  lblT1.Caption = FileNameConvert(lblS1.Caption)
  
End Sub

Function GetFileExt(ByVal strFile As String) As String
'��ȡ�ļ�����չ��
    Dim I As Integer
    Dim intPos As Integer
    intPos = InStrRev(strFile, ".")
    If intPos > 0 Then
        GetFileExt = Mid$(strFile, intPos + 1)
    Else
        GetFileExt = ""
    End If

End Function

Function ChangeFileExt(ByVal strFile As String, ByVal strExt As String) As Boolean
'�޸��ļ���չ��
    Dim I As Integer
    Dim intPos As Integer
    Dim strTmp As String
    intPos = InStrRev(strFile, ".")
    If intPos > 0 Then
        strTmp = Mid$(strFile, 1, intPos) & strExt
        Name strFile As strTmp
    Else
    '��չ��Ϊ��
        Name strFile As strFile & "." & strExt
    End If
    ChangeFileExt = True

End Function

Private Sub cmdChangeExt_Click()
'�޸��ļ���չ��
  Dim strNewExt As String       '����չ��
  Dim I As Integer
  strNewExt = Trim$(txtNewExt.Text)
  If strNewExt = "" Then Exit Sub
  
  For I = 0 To File1.ListCount - 1
    If File1.Selected(I) Then
        If GetFileExt(File1.Path & "\" & File1.List(I)) <> strNewExt Then   '�����չ����������չ������
            Call ChangeFileExt(File1.Path & "\" & File1.List(I), strNewExt)
        End If
    End If
  Next
  File1.Refresh

End Sub
Private Sub cmdGoPath_Click()
On Error GoTo ERR
Dim STRFULLPATH As String
STRFULLPATH = BrowseFolder("���", Me)
If STRFULLPATH <> "" Then File1.Path = STRFULLPATH
ERR:
Exit Sub
End Sub
Private Sub cmdSelect_Click(Index As Integer)
  Dim Ret As Long
  Dim I As Integer
  
  Select Case Index
  Case 0  'ȫѡ
    Ret = SendMessage(File1.hwnd, LB_SETSEL, True, ByVal -1)
    
  Case 1   '��ѡ
    Ret = SendMessage(File1.hwnd, LB_SETSEL, False, ByVal -1)
  Case 2    '��ѡ
    For I = 0 To File1.ListCount - 1
      If File1.Selected(I) Then     '�����ѡ��= True
        File1.Selected(I) = False
      Else
        File1.Selected(I) = True
      End If
    Next
  Case 3
    File1.Refresh
    
  End Select
  
End Sub

Private Sub Combo1_Click()
  Dim temp As Integer
  
  If Len(Combo1.Text) > 4 Then   '�������4λ������Ԥ�ȶ�����ļ���չ��
    temp = InStr(1, Combo1.Text, "(")
    If temp > 0 Then
      File1.Pattern = Mid(Combo1.Text, temp + 1, Len(Combo1.Text) - temp - 1)
'      Debug.Print File1.Pattern
    End If
  Else
    File1.Pattern = Combo1.Text
  End If
  
End Sub

Private Sub Command1_Click()
'�ļ���ת�� ����
  Dim I As Integer
  For I = 0 To File1.ListCount - 1
    If File1.Selected(I) Then   '���ѡ�������
'      Debug.Print File1.Path & "\" & File1.List(i)
      Name File1.Path & "\" & File1.List(I) As File1.Path & "\" & FileNameConvert(File1.List(I))
      
    End If
  Next
  File1.Refresh   'ˢ���ļ��б�
  
End Sub

Private Sub Command2_Click()
'�������� ����
    Dim I As Integer
    mCount = 0            '��ʼ��
    
    For I = 0 To File1.ListCount - 1
        If File1.Selected(I) Then
            'Debug.Print File1.Path & "\" & File1.List(i)
            If Dir$(File1.Path & "\" & BatChangeFileName(File1.List(I))) = "" Then
            '��������ļ������ڣ������
                Name File1.Path & "\" & File1.List(I) As File1.Path & "\" & BatChangeFileName(File1.List(I))
            Else
            '��������ļ��Ѿ����ڣ�����ֱ�Ӹģ�����Ĵ���Ϊֱ������
                Call SHOWWRONG(BatChangeFileName(File1.List(I)) & "�ļ��Ѿ�����!", 0)
            End If
        End If
    Next
    File1.Refresh
  
End Sub


Private Sub cmdRemoveStr_Click()
'ȥ���ļ����Ĳ����ַ�

    Dim strRemove As String     'Ҫȥ�����ַ���
    Dim strTmp As String
    Dim intPos As Integer
    Dim I As Integer
    strRemove = txtRemoveStr.Text
    If strRemove = "" Then Exit Sub     '���û������Ҫȥ�����ַ��� ���˳�
    
    For I = 0 To File1.ListCount - 1
        If File1.Selected(I) Then
            intPos = InStr(File1.List(I), strRemove)
            If intPos > 0 Then
                strTmp = Mid$(File1.List(I), 1, intPos - 1) & Mid$(File1.List(I), intPos + Len(strRemove), Len(File1.List(I)) - intPos - (Len(strRemove) - 1))
                Debug.Print strTmp
                Name File1.Path & "\" & File1.List(I) As File1.Path & "\" & strTmp
            End If
        End If
    Next
    File1.Refresh

End Sub
Private Sub Form_Activate()
Me.BackColor = COLOR_NOR
Dim I As Integer
Me.Cls
For I = 0 To cHECK1.Count - 1
cHECK1(I).M_STYLE = 2
cHECK1(I).SETCOLOR COLOR_NOR, vbWhite
Next
For I = 0 To cmdSelect.Count - 1
cmdSelect(I).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
Next
COMMAND2.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
COMMAND1.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
cmdRemoveStr.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
cmdChangeExt.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
cmdGoPath.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
Dim PBOX As Control
For Each PBOX In Me.Controls
If TypeOf PBOX Is PictureBox Then PBOX.BackColor = Me.BackColor
If TypeOf PBOX Is OptionButton Then PBOX.BackColor = Me.BackColor
Next
Check2(0).BackColor = Me.BackColor
Check2(1).BackColor = Me.BackColor
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
Call PaintPng(App.Path & "\SKIN\N_T.PNG", Me.hdc, 8, 8)
End Sub

Private Sub Form_Load()

cmdGoPath.SETTXT "���"
txtPath.Text = GetInitEntry("NAME", "LASTPATH", App.Path)
Combo1.ListIndex = 0
SCRO.MaxV = PO.Height - PD.ScaleHeight
SCRO.Value = 0
SCRO.LargeChange = 100
cmdSelect(0).SETTXT "ȫѡ"
cmdSelect(1).SETTXT "ȫ��ѡ"
cmdSelect(2).SETTXT "��ѡ"
cmdSelect(3).SETTXT "ˢ��"
COMMAND1.SETTXT "����"
COMMAND2.SETTXT "����"
Me.cmdChangeExt.SETTXT "����"
Me.cmdRemoveStr.SETTXT "����"

cHECK1(0).SETTXT "����"
cHECK1(1).SETTXT "ֻ��"
cHECK1(2).SETTXT "����"
cHECK1(3).SETTXT "ϵͳ"

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub


Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub SCRO_Change()
PO.Top = -SCRO.Value
End Sub
Private Sub txtPath_Change()
lRet = SetInitEntry("NAME", "LASTPATH", txtPath.Text)
File1.Path = txtPath.Text
End Sub

Private Sub x1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = False
X2.Visible = True
End Sub
Private Sub x2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
X2.Visible = False
X3.Visible = True
End If
End Sub
Private Sub x3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X3.Visible = False
X1.Visible = True
If X3.Visible = False Then Unload Me
End Sub


'ת����
Private Sub Option1_Click(Index As Integer)
    lblT1.Caption = FileNameConvert(lblS1.Caption)
    
End Sub

'���Ҳ����
Function StrCls(sSource As String, sSearch As String) As String
  Dim I As Integer, res As String
  res = sSource
  Do While InStr(res, sSearch)
    I = InStr(res, sSearch)
    res = Left(res, I - 1) & Mid(res, I + Len(sSearch))
  Loop
  StrCls = res
  
End Function
'���Ҳ��滻
Function StrRel(sSource As String, sSearch As String, rel As String) As String
  Dim I As Integer, res As String
  res = sSource
  Do While InStr(res, sSearch)
    I = InStr(res, sSearch)
    res = Left(res, I - 1) & rel & Mid(res, I + Len(sSearch))
  Loop
  StrRel = res
End Function

Private Sub Option2_Click(Index As Integer)
  '�ı��� �ո��� ��ʽ�����´���
    lblT1.Caption = FileNameConvert(lblS1.Caption)
    
End Sub

'�Զ����ı�λ��
Private Sub Option3_Click(Index As Integer)
    mCount = 0
    lblTarget(0).Caption = BatChangeFileName(lblSource(0).Caption)
    lblTarget(1).Caption = BatChangeFileName(lblSource(1).Caption)
    

End Sub


Private Function FileNameConvert(OldFileName As String) As String
'�ļ���ת�� ����

  Dim I As Integer
  
  '������չ�� ���λ��
  I = InStrRev(OldFileName, ".")        '���ַ�����ĩβ����
  If I = 0 Then I = Len(OldFileName) + 1  '����ļ���û��Dot�Ļ�����ֵΪ�ļ�����+1
  
  Select Case True
  Case Option1(0)  '��д
    If Check2(0).Value And Check2(1).Value Then   'ǰ��׺
      FileNameConvert = UCase$(OldFileName)
    ElseIf Check2(0).Value Then   'ǰ׺
      FileNameConvert = UCase$(Left$(OldFileName, I - 1)) & "." & Mid$(OldFileName, I + 1, Len(OldFileName) - I)
    ElseIf Check2(1).Value Then   '��׺
      FileNameConvert = Left$(OldFileName, I - 1) & "." & UCase$(Right$(OldFileName, Len(OldFileName) - I))
    End If
    
  Case Option1(1)  'Сд
    If Check2(0).Value And Check2(1).Value Then   'ǰ��׺
      FileNameConvert = LCase$(OldFileName)
    ElseIf Check2(0).Value Then   'ǰ׺
      FileNameConvert = LCase$(Left$(OldFileName, I - 1)) & "." & Mid$(OldFileName, I + 1, Len(OldFileName) - I)
    ElseIf Check2(1).Value Then   '��׺
      FileNameConvert = Left$(OldFileName, I - 1) & "." & LCase$(Right$(OldFileName, Len(OldFileName) - I))
    End If
  
  Case Option1(2)   '����ĸ��д
    If Check2(0).Value And Check2(1).Value Then   'ǰ��׺
      FileNameConvert = UCase$(Left$(OldFileName, 1)) & _
            LCase$(Right$(OldFileName, Len(OldFileName) - 1))
    ElseIf Check2(0).Value Then   'ǰ׺
      FileNameConvert = UCase$(Left$(OldFileName, 1)) & LCase$(Mid$(OldFileName, 2, I - 1 - 1)) _
            & "." & Right$(OldFileName, Len(OldFileName) - I)
    ElseIf Check2(1).Value Then   '��׺
      FileNameConvert = Left$(OldFileName, I - 1) & "." & UCase$(Mid$(OldFileName, I + 1, 1)) & LCase$(Right$(OldFileName, Len(OldFileName) - I - 1))
    End If
    
  End Select
  
  If Option2(0).Value Then
    FileNameConvert = StrCls(FileNameConvert, " ")
  ElseIf Option2(1).Value Then
    FileNameConvert = StrRel(FileNameConvert, " ", "_")
  End If

End Function

Private Function BatChangeFileName(OldFileName As String) As String
  Dim s As Integer
  Dim I As Integer
  Dim strTmp As String
  
  strTmp = Trim(Text1.Text)
  s = CInt(txtStep.Text)
  mCount = mCount + s
  
  I = InStrRev(OldFileName, ".")  '���ҵ��λ��
  If I = 0 Then I = Len(OldFileName) + 1
  
  Select Case True
  Case Option3(0)  '�滻ǰ׺
    BatChangeFileName = strTmp & mCount & Right$(OldFileName, Len(OldFileName) - I + 1)  '��1�ǰ��Ǹ���ӽ���
    
  Case Option3(1)  '�滻��׺
    BatChangeFileName = VBA.Left$(OldFileName, I) & strTmp & mCount
    
  Case Option3(2)  '��ǰ׺��ǰ��
    BatChangeFileName = strTmp & mCount & OldFileName
    
  Case Option3(3)  '��ǰ׺�ĺ���
    BatChangeFileName = VBA.Left$(OldFileName, I - 1) & strTmp & mCount & Right$(OldFileName, Len(OldFileName) - I + 1)
    
  Case Option3(4)  '�ں�׺��ǰ��
    BatChangeFileName = VBA.Left$(OldFileName, I) & strTmp & mCount & Right$(OldFileName, Len(OldFileName) - I)
    
  Case Option3(5)  '�ں�׺�ĺ���
    BatChangeFileName = OldFileName & strTmp & mCount
    
  End Select

End Function

Private Sub Text1_Change()
    mCount = 0
    lblTarget(0).Caption = BatChangeFileName(lblSource(0).Caption)
    lblTarget(1).Caption = BatChangeFileName(lblSource(1).Caption)
    

End Sub

Private Sub txtPath_DblClick()
    txtPath.SelStart = 0
    txtPath.SelLength = Len(txtPath.Text)
    
End Sub

Private Sub txtPath_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then       '��·���򰴻س�����ִ��ת������
        KeyAscii = 0
        Call cmdGoPath_Click
        
    End If
    
End Sub

Private Sub txtStep_Change()
    mCount = 0
    lblTarget(0).Caption = BatChangeFileName(lblSource(0).Caption)
    lblTarget(1).Caption = BatChangeFileName(lblSource(1).Caption)

End Sub


