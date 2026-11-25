VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SocketOne Client"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8281
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Server Connect"
      TabPicture(0)   =   "frmMain.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frm1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Buddies && Chat"
      TabPicture(1)   =   "frmMain.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstBuddies"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Misc"
      TabPicture(2)   =   "frmMain.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.ListBox lstBuddies 
         Height          =   2985
         Left            =   -74760
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin VB.Frame frm1 
         Caption         =   "SocketOne Information:"
         Height          =   1695
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   4695
         Begin VB.TextBox txtStatus 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   1680
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   960
            Width           =   2895
         End
         Begin VB.TextBox txtOutPort 
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   1680
            TabIndex        =   6
            Text            =   "52876"
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox txtOutHost 
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   1680
            TabIndex        =   5
            Text            =   "192.168.1.133"
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "SocketOne Port:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "SocketOne Server:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Status:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "SocketOne Client Interaction:"
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   2280
         Width           =   4695
         Begin VB.CheckBox chkNew 
            Caption         =   "Check here if you are a new user seeking an account."
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1560
            Width           =   4215
         End
         Begin VB.TextBox txtPassword2 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1680
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   16
            Top             =   1200
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.TextBox txtPassword 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1680
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   15
            Top             =   840
            Width           =   2895
         End
         Begin VB.ComboBox cboLogin 
            Height          =   315
            Left            =   1680
            TabIndex        =   14
            Top             =   360
            Width           =   2895
         End
         Begin VB.CommandButton cmdConnect 
            Caption         =   "LogOn SocketOne"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   1920
            Width           =   4455
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "LogOff SocketOne"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   1920
            Visible         =   0   'False
            Width           =   4455
         End
         Begin VB.Label lblPassword2 
            Caption         =   "Re-Enter Password:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1200
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Password:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Login:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   495
         End
      End
   End
   Begin MSWinsockLib.Winsock Wsck 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   52876
   End
   Begin VB.Timer TimerControl 
      Interval        =   500
      Left            =   0
      Top             =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkNew_Click()
If chkNew.Value = Checked Then
    lblPassword2.Visible = True
    txtPassword2.Visible = True
ElseIf chkNew.Value = Unchecked Then
    lblPassword2.Visible = False
    txtPassword2.Visible = False
End If
End Sub

Private Sub cmdClose_Click()
Wsck.Close
CloseEvent
End Sub

Private Sub cmdConnect_Click()
SendStatus "Connection to SocketOne Server."
Wsck.RemoteHost = txtOutHost.Text
Wsck.RemotePort = Val(txtOutPort.Text)
ConnectEvent
DoEvents

Wsck.Connect
TimerControl.Tag = Timer + 30
End Sub

Private Sub Form_Load()
SendStatus "Ready"
End Sub

Private Sub TimerControl_Timer()
If Not (TimerControl.Tag = "") Then
    If Timer > Val(TimerControl.Tag) Then
        Call cmdClose_Click
        MsgBox "Connection Attempted timed out."
        TimerControl.Tag = ""
    End If
End If
End Sub

Private Sub Wsck_Close()
SendStatus "Disconnected."
Wsck.Close
CloseEvent
End Sub

Private Sub Wsck_Connect()
SendStatus "Connection Established."
cmdClose.Visible = True
DoEvents
'LoginAttempt
End Sub

Private Sub SendStatus(Sts As String)
    txtStatus.Text = txtStatus.Text & Sts & vbCrLf
End Sub

Private Sub Wsck_DataArrival(ByVal bytesTotal As Long)
Dim InData, SCode As String
Wsck.GetData InData, vbString
SCode = CodeCheck(InData)
InData = CodeLess(InData)

Select Case SCode
Case ServerMessage
    MsgBox InData
Case Failure
    Call cmdClose_Click
Case RequestLogin
    SocketSend (LoginCode & UCase(Trim(cboLogin.Text)) + Chr(0) & UCase(Trim(txtPassword.Text)))
Case Else
    'In Case
End Select

End Sub

Private Sub ConnectEvent()
cmdConnect.Visible = False
cboLogin.Enabled = False
txtPassword.Enabled = False
txtPassword2.Enabled = False
TimerControl.Tag = ""
End Sub

Private Sub CloseEvent()
cmdConnect.Visible = True
cmdClose.Visible = False
cboLogin.Enabled = True
txtPassword.Enabled = True
txtPassword2.Enabled = True
TimerControl.Tag = ""
End Sub
