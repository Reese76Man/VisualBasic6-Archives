VERSION 5.00
Begin VB.Form frmMainMenu 
   BorderStyle     =   0  'None
   Caption         =   "MCLE"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGroups 
      Caption         =   "Groups"
      Height          =   975
      Left            =   3720
      Picture         =   "frmMainMenu.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdReports 
      Caption         =   "Reports"
      Height          =   975
      Left            =   1800
      Picture         =   "frmMainMenu.frx":044A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Enter DBE General section."
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdSession 
      Caption         =   "Retirees"
      Height          =   975
      Left            =   120
      Picture         =   "frmMainMenu.frx":088C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Enter Wholesale section."
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtCurrentUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Currently Logged in Network User."
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "SQL Server Datestamp."
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit FRP"
      Height          =   975
      Left            =   5400
      Picture         =   "frmMainMenu.frx":0CCE
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Exit BATS to Windows."
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Current Local Time"
      Top             =   2160
      Width           =   975
   End
   Begin VB.Timer Timer 
      Interval        =   450
      Left            =   0
      Top             =   0
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   6600
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   6600
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3840
   End
   Begin VB.Line Line1 
      X1              =   6600
      X2              =   6600
      Y1              =   0
      Y2              =   3840
   End
   Begin VB.Image IMGFatTuesday 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   2025
      Left            =   0
      Picture         =   "frmMainMenu.frx":1110
      Stretch         =   -1  'True
      ToolTipText     =   "Fat Tuesday...mmm mmm good shit!"
      Top             =   0
      Width           =   2085
   End
   Begin VB.Label lblACTS 
      Alignment       =   2  'Center
      Caption         =   "F R P"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   2040
      TabIndex        =   11
      ToolTipText     =   "Main Menu"
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label lblUser 
      Caption         =   "User:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblDate 
      Caption         =   "Date:"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   2160
      Width           =   495
   End
   Begin VB.Line UnderLine 
      X1              =   0
      X2              =   6600
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Retirement Tracking System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      ToolTipText     =   "Main Menu"
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Line Line15 
      X1              =   0
      X2              =   6600
      Y1              =   2010
      Y2              =   2010
   End
   Begin VB.Label Label2 
      Caption         =   "Time:"
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   2160
      Width           =   495
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
If UserClick = True Then
    StartRoutine Me
        If MsgBox("Exit FRP Completely?", vbYesNo, "Exit") = vbYes Then
            ExitProgram
        End If
        MDIMain.MNUMainMenu.Enabled = True
    EndRoutine Me
    Unload frmMainMenu
End If
End Sub

Private Sub cmdGroups_Click()
If UserClick = True Then OpenWindow frmGroup, ""
End Sub

Private Sub cmdReports_Click()
If UserClick = True Then
    MDIMain.MNUReports.Enabled = False
    LoadForm frmEntryReports
End If
End Sub

Private Sub cmdSession_Click()
If UserClick = True Then OpenWindow frmEntry, ""
End Sub

Private Sub Form_Load()
StartRoutine Me
    txtDate.Text = SQLDate
    txtTime.Text = Time$
    txtCurrentUser.Text = PCName
EndRoutine Me
End Sub

Private Sub Timer_Timer()
If Not Left(txtTime.Text, 2) = Left(Time$, 2) Then
    GetSQLDate
    txtDate.Text = SQLDate
End If
txtTime.Text = Time$
End Sub
