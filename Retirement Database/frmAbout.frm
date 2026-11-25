VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programmed by Corey Michael Badeaux"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtBarCode 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1920
      Width           =   4815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4920
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label lblversion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VERSION"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.Label LBLMain 
      Alignment       =   2  'Center
      Caption         =   "FRP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Version 1.1"
      Top             =   0
      Width           =   4935
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4920
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblLADB 
      Alignment       =   2  'Center
      Caption         =   "Copyright © 2006 Kinyodas Software"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label lblTeam 
      Alignment       =   2  'Center
      Caption         =   "Programming && Development:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label LTeam2 
      Alignment       =   2  'Center
      Caption         =   "Corey Badeaux"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   4815
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MCLE Development: Corey Badeaux.
' Programmed by Corey Michael Badeaux
Option Explicit

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Activate()
cmdOK.SetFocus
End Sub

Private Sub Form_Load() 'A product so good, we put our name on it!
  lblversion.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.MNUAbout.Enabled = True
End Sub
