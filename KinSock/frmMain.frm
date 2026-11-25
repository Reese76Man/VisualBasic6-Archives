VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KinSock v1.0 - KinYodas Software"
   ClientHeight    =   10245
   ClientLeft      =   570
   ClientTop       =   585
   ClientWidth     =   3915
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10245
   ScaleWidth      =   3915
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   9720
      Width           =   3735
   End
   Begin VB.PictureBox PicBox 
      Height          =   9015
      Left            =   120
      ScaleHeight     =   8955
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   600
      Width           =   3735
   End
   Begin MSWinsockLib.Winsock WSK 
      Index           =   0
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
SockSetup
End Sub


Private Sub SockSetup()
Dim t As Long

PicBox.Print "Starting Init..."
PicBox.Print "Host:" & WSK(0).LocalHostName
PicBox.Print "IP:" & WSK(0).LocalIP
PicBox.Print "Assigning Listen Ports:..."
For t = 1 To 1000
    't = 0
    Load WSK(t)
    WSK(t).LocalPort = t
    On Error Resume Next
    WSK(t).Listen
    If Err.Number = 10048 Then
        PicBox.Print "Cannot Listen to Port in Use:" & t
        Err.Number = 0
    End If
Next t
PicBox.Print "Done!"
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim t
For t = 1 To 1000
    If WSK(t).State = 2 Then WSK(t).Close
Next t
End Sub

Private Sub WSK_ConnectionRequest(Index As Integer, ByVal requestID As Long)
PicBox.Print "Incoming Request on Port:" & Index
End Sub


