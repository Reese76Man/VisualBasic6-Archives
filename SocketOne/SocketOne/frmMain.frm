VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SocketOne"
   ClientHeight    =   8505
   ClientLeft      =   2670
   ClientTop       =   1680
   ClientWidth     =   10050
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   9
      TabHeight       =   520
      TabCaption(0)   =   "Main"
      TabPicture(0)   =   "frmMain.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmStats"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "SocketOne User"
      TabPicture(1)   =   "frmMain.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Label4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Statistics"
      TabPicture(2)   =   "frmMain.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label5"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Chat Room Monitor"
      TabPicture(3)   =   "frmMain.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "File Transfer Monitor"
      TabPicture(4)   =   "frmMain.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Server Settings"
      TabPicture(5)   =   "frmMain.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Games Monitor"
      TabPicture(6)   =   "frmMain.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Individual Chat Monitor"
      TabPicture(7)   =   "frmMain.frx":0506
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "DataBase Monitor"
      TabPicture(8)   =   "frmMain.frx":0522
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      Begin VB.Frame Frame2 
         Caption         =   "Server Status"
         Height          =   6255
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   3255
         Begin VB.TextBox txtInIp 
            Height          =   285
            Left            =   1080
            TabIndex        =   15
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtInPort 
            Height          =   285
            Left            =   1080
            TabIndex        =   14
            Text            =   "52876"
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtInHost 
            Height          =   285
            Left            =   1080
            TabIndex        =   12
            Top             =   960
            Width           =   2055
         End
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
            Height          =   4725
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   1440
            Width           =   3015
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   3120
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Label lbl3 
            Caption         =   "Listen Port:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lbl2 
            Caption         =   "Server IP:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lbl4 
            Caption         =   "Server Host:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame frmStats 
         Caption         =   "User Information / Options"
         Height          =   6255
         Left            =   5880
         TabIndex        =   3
         Top             =   1560
         Width           =   3495
         Begin VB.CommandButton cmdChat 
            Caption         =   "Chat User"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   4920
            Width           =   975
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit"
            Height          =   255
            Left            =   1200
            TabIndex        =   47
            Top             =   4920
            Width           =   1095
         End
         Begin VB.CommandButton cmdKick 
            Caption         =   "Kick User"
            Height          =   255
            Left            =   2400
            TabIndex        =   46
            Top             =   4920
            Width           =   975
         End
         Begin VB.TextBox txtCreated 
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   4440
            Width           =   1095
         End
         Begin VB.TextBox txtCountry 
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   3000
            Width           =   615
         End
         Begin VB.TextBox txtSex 
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   4080
            Width           =   615
         End
         Begin VB.TextBox txtBirthDay 
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   4080
            Width           =   1095
         End
         Begin VB.TextBox txtPhone 
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   3720
            Width           =   2415
         End
         Begin VB.TextBox txtEmail 
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   3360
            Width           =   2415
         End
         Begin VB.TextBox txtMiddleName 
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   1560
            Width           =   2415
         End
         Begin VB.TextBox txtLastName 
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   1920
            Width           =   2415
         End
         Begin VB.TextBox txtAddress 
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   2280
            Width           =   2415
         End
         Begin VB.TextBox txtCity 
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   2640
            Width           =   2415
         End
         Begin VB.TextBox txtState 
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   3000
            Width           =   1095
         End
         Begin VB.TextBox txtFirstName 
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1200
            Width           =   2415
         End
         Begin VB.TextBox txtPass 
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox txtLogin 
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label17 
            Caption         =   "Created:"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   4440
            Width           =   855
         End
         Begin VB.Line Line3 
            X1              =   120
            X2              =   3360
            Y1              =   4800
            Y2              =   4800
         End
         Begin VB.Label Label16 
            Caption         =   "BirthDay:"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   4080
            Width           =   735
         End
         Begin VB.Label Label15 
            Caption         =   "Country:"
            Height          =   255
            Left            =   2160
            TabIndex        =   37
            Top             =   3000
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "Phone:"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   3720
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Email:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   3360
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Sex:"
            Height          =   255
            Left            =   2160
            TabIndex        =   34
            Top             =   4080
            Width           =   375
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   3360
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label11 
            Caption         =   "First:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Middle:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "Last:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Address:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "City:"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "State:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   3000
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Password:"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Login:"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Connected Sockets"
         Height          =   6255
         Left            =   3480
         TabIndex        =   8
         Top             =   1560
         Width           =   2295
         Begin VB.ListBox LSTConnects 
            Height          =   5910
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Buddies"
         Height          =   6255
         Left            =   -74880
         TabIndex        =   20
         Top             =   1560
         Width           =   1815
         Begin VB.ListBox lstBuddies 
            Height          =   5910
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Server Status "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -74880
         TabIndex        =   19
         Top             =   1080
         Width           =   9255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SocketOne User Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -74880
         TabIndex        =   18
         Top             =   1080
         Width           =   9255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Server Status "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   9255
      End
   End
   Begin MSWinsockLib.Winsock WSCK 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SocketOne Virtual Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AddStatus(Sts As String)
txtStatus.Text = txtStatus.Text + Sts + vbCrLf
End Sub

Private Sub WSKSetup()
WSCK(0).Protocol = sckTCPProtocol            'Using TCP/IP
WSCK(0).LocalPort = Val(txtInPort.Text)      'Using Set Port (Default: 52876)
WSCK(0).Bind WSCK(0).LocalPort, txtInIp.Text 'Binds Port/IP to ensure Listen
WSCK(0).Listen                               'Sets Socket to Listen ONLY
AddStatus ("Listening Mode...")    'Adds status of Listening Mode to status.

LSTConnects.AddItem "SocketOne", 0 'Adds the console to the user list.
LSTConnects.ListIndex = 0          'Sets to Highight console entry
SocketLogin(0) = UserLogin("SocketOne", "Peace")
LSTConnects_Click                  'Validates and Displays console information
End Sub

Private Sub Form_Load()
Set RS = New ADODB.Recordset            'Sets new Recordset Object
DE.CN.Open                              'Opens Connection to DataBase
RS.Open "Users", DE.CN, adOpenStatic, adLockOptimistic  'Opens Users Table

txtInIp.Text = WSCK(0).LocalIP          'Gets Local IP on default Network
txtInHost.Text = WSCK(0).LocalHostName  'Gets Local HostName on default Network
WSKSetup                                'Initiates Winsock Setup Routines
End Sub

Private Sub Form_Unload(Cancel As Integer)
RS.Close
DE.CN.Close
End Sub

Private Sub LSTConnects_Click()
If UserStats(LSTConnects.ItemData(LSTConnects.ListIndex)) = True Then
    txtLogin.Text = "" & Trim(RS!UserName)
    txtPass.Text = "" & Trim(RS!UserPass)
    txtFirstName.Text = "" & Trim(RS!FirstName)
    txtMiddleName.Text = "" & Trim(RS!MiddleName)
    txtLastName.Text = "" & Trim(RS!LastName)
    txtAddress.Text = "" & Trim(RS!Address)
    txtCity.Text = "" & Trim(RS!City)
    txtState.Text = "" & Trim(RS!State)
    txtCountry.Text = "" & Trim(RS!Country)
    txtEmail.Text = "" & Trim(RS!Email)
    txtSex.Text = "" & Trim(RS!Sex)
    txtBirthDay.Text = "" & Trim(RS!BirthDate)
    txtPhone.Text = "" & Trim(RS!Phone)
    txtCreated.Text = "" & Trim(RS!UserFirstDate)
End If
End Sub

Private Sub WSCK_Close(Index As Integer)
AddStatus ("Closing from: " & WSCK(Index).RemoteHostIP)
ManualClose Index          'Closes and Cleans up Socket
End Sub

Private Sub WSCK_ConnectionRequest(Index As Integer, ByVal RequestID As Long)
Dim P, UnusedSocket, SocketNum As Integer
AddStatus ("Connection request from: " & WSCK(Index).RemoteHostIP)

If SocketCount > 32000 Then
    For SocketNum = 1 To SocketCount
        If SocketOpen(SocketNum) = False Then UnusedSocket = SocketNum
    Next SocketNum
End If

If UnusedSocket = 0 Then
    SocketCount = SocketCount + 1
    SocketNum = SocketCount
Else
    SocketNum = UnusedSocket
End If

Load WSCK(SocketNum)
WSCK(SocketNum).Accept RequestID
LSTConnects.AddItem (WSCK(SocketNum).RemoteHostIP), P
LSTConnects.ItemData(P) = SocketNum
SocketSend SocketNum, RequestLogin
SocketOpen(SocketNum) = True
End Sub


Private Sub WSCK_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim InData, SCode As String
Dim IDex As Long
WSCK(Index).GetData InData, vbString
SCode = CodeCheck(InData)
InData = CodeLess(InData)

If SocketLogin(Index) = 0 Then      'All Actions made Before login
    Select Case SCode
        Case LoginCode
            IDex = UserLogin(Login2User(InData), Login2Pass(InData))
            If IDex <= 0 Then
                ServerSend Index, "Username or Password invalid!"
                SocketSend Index, Failure
            End If
            If IDex > 0 Then
                SocketLogin(Index) = IDex
                SocketSend Index, Success
                UserConnected Index, IDex
            End If
    End Select
Else                                'All Actions made After login
End If
End Sub

