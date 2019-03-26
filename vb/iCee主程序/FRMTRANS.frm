VERSION 5.00
Begin VB.Form FRMTRAN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "远程协助"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11925
   Icon            =   "FRMTRANS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   529
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PO 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00231C09&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   11925
      TabIndex        =   1
      Top             =   0
      Width           =   11925
      Begin ICEE.ICEE_COMMAND COK 
         Height          =   615
         Left            =   2760
         TabIndex        =   9
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
      End
      Begin VB.OptionButton OptIsServer 
         BackColor       =   &H00231C09&
         Caption         =   "客户机(&C)"
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   0
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   6
         Top             =   120
         Width           =   1200
      End
      Begin VB.OptionButton OptIsServer 
         BackColor       =   &H00231C09&
         Caption         =   "服务器(&S)"
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   1
         Left            =   1440
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   5
         Top             =   120
         Width           =   1200
      End
      Begin VB.TextBox TxtIP 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   960
         TabIndex        =   4
         Text            =   "127.0.0.1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox TxtPort 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   960
         TabIndex        =   3
         Text            =   "1503"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TxtPort 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   1
         Left            =   1800
         TabIndex        =   2
         Text            =   "1312"
         Top             =   960
         Width           =   855
      End
      Begin ICEE.ICEE_COMMAND CUN 
         Height          =   615
         Left            =   4560
         TabIndex        =   10
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
      End
      Begin VB.Label LblIP 
         BackStyle       =   0  'Transparent
         Caption         =   "IP地址："
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   720
      End
      Begin VB.Label LblPort 
         BackStyle       =   0  'Transparent
         Caption         =   "端口号："
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   720
      End
   End
   Begin ICEE.ICEE_NET NetGIFTran1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   11895
      _ExtentX        =   7858
      _ExtentY        =   1931
   End
End
Attribute VB_Name = "FRMTRAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Option Explicit

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Const HALFTONE = 4
Private mOldTime As Currency

Private Sub Form_Load()
COK.SETTXT "连    接"
CUN.SETTXT "关    闭"
End Sub

Private Sub Form_Paint()
    NetGIFTran1.Draw Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
Private Sub Form_Unload(Cancel As Integer)
    NetGIFTran1.CloseConnect
End Sub
Private Sub NetGIFTran1_CloseConnect()
    Call UpdataUI
End Sub
Private Sub NetGIFTran1_OnPictureArrival()
    Dim CurTime As Currency
    
    CurTime = GetCurTime()
    Me.Caption = App.Title & ": " & vbTab & "间隔:" & Format$(CurTime - mOldTime, "###,###,###,##0.0000") & "ms"
    mOldTime = CurTime
    
    Me.Refresh
    
End Sub

Private Sub NetGIFTran1_OnQueryPicture()
    'Debug.Print "NetGIFTran1_OnQueryPicture"
    
    Static OldTime As Currency
    Dim CurTime As Currency
    Const SampTimeDis = 1000 '采样间隔(ms)
    
    If NetGIFTran1.CurClients > 1 Then '只有客户数大于1时才控制采样间隔
        CurTime = GetCurTime()
        If OldTime + SampTimeDis > CurTime Then '未到时间
            Exit Sub
        End If
        OldTime = CurTime
    End If
    
    Dim hDCScr As Long
    
    hDCScr = GetDC(0)
    If hDCScr Then
        Me.MousePointer = vbHourglass
        Call NetGIFTran1.SetBitmap(hDCScr, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY)
        Me.MousePointer = vbDefault
        
        Call ReleaseDC(0, hDCScr)
        
        Me.Refresh
        
    End If
    
End Sub
'更新界面
Public Sub UpdataUI()
    Dim fConnect As Boolean
    
    With Me.NetGIFTran1
        If sckClosing = .State Then .CloseConnect
        fConnect = (.State <> sckClosed)
        
        OptIsServer(0).Enabled = Not fConnect
        OptIsServer(1).Enabled = Not fConnect
        TxtIP.Enabled = Not fConnect
        TxtPort(0).Enabled = Not fConnect
        TxtPort(1).Enabled = Not fConnect
        
        Dim TempStr As String
        If fConnect Then
            TempStr = .RemoteHostIP
            If Len(TempStr) = 0 Then TempStr = .RemoteHost
        Else
            TempStr = .LocalIP
        End If
        If Len(TempStr) Then TxtIP.Text = TempStr
        
        OptIsServer(0).Value = Not .IsServer
        OptIsServer(1).Value = .IsServer
        'Debug.Print .IsServer
        
        If fConnect Then
            If .IsServer Then
                TxtPort(0).Text = CStr(.LocalPort)
                TxtPort(1).Text = CStr(.RemotePort)
            Else
                TxtPort(0).Text = CStr(.RemotePort)
                TxtPort(1).Text = CStr(.LocalPort)
            End If
            
        End If

    End With
    
End Sub

Private Sub CUN_Click()
    With NetGIFTran1
        If .State <> sckClosed Then '已连结
            Call .CloseConnect
            
        End If
        
    End With
    
    Call UpdataUI
    
End Sub

Private Sub COK_Click()
    With NetGIFTran1
        If .State = sckClosed Then '未连结
            .IsServer = OptIsServer(1).Value
            If .IsServer Then
                .LocalPort = Val(TxtPort(0).Text)
                .RemotePort = Val(TxtPort(1).Text)
            Else
                .RemotePort = Val(TxtPort(0).Text)
                .LocalPort = 0 'Val(TxtPort(1).Text)
                .RemoteHost = Trim(TxtIP.Text)
            End If
            
            If .Connect() = False Then Call SHOWWRONG("无法连结!", 2)
            
        End If
        
    End With
    
    Call UpdataUI
    
End Sub

Private Sub Form_Activate()
    Call UpdataUI
    
End Sub

Private Sub TxtIP_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKey0 To vbKey9
    Case vbKeyBack
    Case 46 'Asc(".")=46
    Case Else
        KeyAscii = 0
    End Select
    
End Sub

Private Sub TxtPort_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKey0 To vbKey9
    Case vbKeyBack
    Case Else
        KeyAscii = 0
    End Select
    
End Sub


