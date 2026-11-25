VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programmed by Corey Michael Badeaux"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4920
      Y1              =   1920
      Y2              =   1920
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
      Caption         =   "East Tree"
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
      Caption         =   "Copyright © 2014 - East Tree Solutions LLC"
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
      Caption         =   "Corey Badeaux "
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4815
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmed & Developed by Corey Michael Badeaux
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


