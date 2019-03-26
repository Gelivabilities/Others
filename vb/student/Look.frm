VERSION 5.00
Begin VB.Form frmView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "你的信息如下："
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   5310
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Frame fraInfo 
      Enabled         =   0   'False
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.TextBox txtSerial 
         DataField       =   "Serial"
         DataMember      =   "Student"
         DataSource      =   "DataEnv"
         Height          =   330
         Left            =   1080
         TabIndex        =   16
         Top             =   315
         Width           =   1215
      End
      Begin VB.TextBox txtName 
         DataField       =   "Name"
         DataMember      =   "Student"
         DataSource      =   "DataEnv"
         Height          =   300
         Left            =   1080
         TabIndex        =   7
         Top             =   780
         Width           =   1215
      End
      Begin VB.TextBox txtBirthday 
         DataField       =   "Birthday"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         DataMember      =   "Student"
         DataSource      =   "DataEnv"
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   1665
         Width           =   1200
      End
      Begin VB.TextBox txtAddress 
         DataField       =   "Address"
         DataMember      =   "Student"
         DataSource      =   "DataEnv"
         Height          =   885
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2520
         Width           =   3720
      End
      Begin VB.TextBox txtResume 
         DataField       =   "resume"
         DataMember      =   "Student"
         DataSource      =   "DataEnv"
         Height          =   1005
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   3960
         Width           =   3720
      End
      Begin VB.TextBox txtTelephone 
         DataField       =   "tel"
         DataMember      =   "Student"
         DataSource      =   "DataEnv"
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   3540
         Width           =   2400
      End
      Begin VB.TextBox txtClass 
         DataField       =   "class"
         DataMember      =   "Student"
         DataSource      =   "DataEnv"
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   1215
         Width           =   1215
      End
      Begin VB.TextBox txtFalse 
         DataField       =   "sex"
         DataMember      =   "Student"
         DataSource      =   "DataEnv"
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         Top             =   2085
         Width           =   735
      End
      Begin VB.Image imgPhoto 
         DataField       =   "Photo"
         DataMember      =   "Student"
         DataSource      =   "DataEnv"
         Height          =   2175
         Left            =   2520
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "学号:"
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   17
         Top             =   390
         Width           =   450
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "姓名:"
         Height          =   180
         Index           =   1
         Left            =   480
         TabIndex        =   14
         Top             =   840
         Width           =   450
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "班级:"
         Height          =   180
         Index           =   11
         Left            =   480
         TabIndex        =   13
         Top             =   1282
         Width           =   450
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "出生日期:"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1717
         Width           =   810
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "性别:"
         Height          =   180
         Index           =   5
         Left            =   480
         TabIndex        =   11
         Top             =   2145
         Width           =   450
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "地址:"
         Height          =   180
         Index           =   6
         Left            =   480
         TabIndex        =   10
         Top             =   2872
         Width           =   450
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "简历:"
         Height          =   300
         Index           =   8
         Left            =   480
         TabIndex        =   9
         Top             =   4312
         Width           =   450
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "电话:"
         Height          =   180
         Index           =   7
         Left            =   480
         TabIndex        =   8
         Top             =   3592
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    ''根据当前登录的用户在DataEnv.rsStudent中查找到对应的记录
   DataEnv.rsStudent.Find "serial = '" & MDIMain.msUserName & "'"
End Sub

