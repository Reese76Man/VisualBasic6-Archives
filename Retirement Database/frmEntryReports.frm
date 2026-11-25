VERSION 5.00
Begin VB.Form frmEntryReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entry Reports"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   Icon            =   "frmEntryReports.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   8370
   Begin VB.CommandButton cmdReport 
      Caption         =   "Create Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7320
      Picture         =   "frmEntryReports.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Preview & Print selected recipe."
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7320
      Picture         =   "frmEntryReports.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.Frame Frame5 
      Caption         =   "Retiree Reports"
      Height          =   2175
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7215
      Begin VB.Frame FrameAll 
         Height          =   420
         Left            =   2880
         TabIndex        =   5
         Top             =   115
         Width           =   4215
         Begin VB.OptionButton OptDeceased 
            Caption         =   "Deceased"
            Height          =   195
            Left            =   3000
            TabIndex        =   9
            Top             =   145
            Width           =   1095
         End
         Begin VB.OptionButton OPTNonActive 
            Caption         =   "Non-Active"
            Height          =   195
            Left            =   1680
            TabIndex        =   8
            Top             =   145
            Width           =   1095
         End
         Begin VB.OptionButton OptActive 
            Caption         =   "Active"
            Height          =   195
            Left            =   720
            TabIndex        =   7
            Top             =   145
            Width           =   855
         End
         Begin VB.OptionButton OptAll2 
            Caption         =   "All"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.OptionButton OptDues 
         Caption         =   "All Retirees - With Dues Paid Information"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   3495
      End
      Begin VB.OptionButton OptAll 
         Caption         =   "All Retirees - Contact Information"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   265
         Value           =   -1  'True
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmEntryReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdReport_Click()
If UserClick = True Then
    UserClick = False
    
    If OptAll.Value = True Then
        If OptAll2.Value = True Then
        ElseIf OptActive.Value = True Then
        ElseIf OPTNonActive.Value = True Then
        ElseIf OptDeceased.Value = True Then
        End If
    
        DE.EntryReport1 OptAll2.Tag, OptActive.Tag, OPTNonActive.Tag, OptDeceased.Tag
        frmReport.ViewReport "EntryReport", DE.rsEntryReport1
        LoadForm frmReport, 1
        DE.rsEntryReport1.Close
    End If
    
    UserClick = True
End If
End Sub

Private Sub Form_Load()
cboGroup.Tag = FillTemplate(cboGroup, "MembershipGroup", "Description", "MembershipGroupID")
cboGroup.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.MNUReports.Enabled = True
End Sub

Private Sub OptAll_Click()
If OptAll.Value = True Then
    FrameAll.Enabled = True
Else
    FrameAll.Enabled = False
End If
End Sub

Private Sub OptDues_Click()
OptAll_Click
End Sub
