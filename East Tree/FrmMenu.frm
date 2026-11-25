VERSION 5.00
Begin VB.Form FrmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "East Tree LLC"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "FrmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   326
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4935
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   3600
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Timer Timer 
         Interval        =   900
         Left            =   0
         Top             =   0
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
         Left            =   2160
         TabIndex        =   13
         ToolTipText     =   "Version 1.1"
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label lblversion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "VERSION"
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
      Begin VB.Image Image 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1890
         Left            =   120
         Picture         =   "FrmMenu.frx":0442
         ToolTipText     =   "Where's my fucking logo, Kenny?"
         Top             =   240
         Width           =   1830
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "MENU OPTIONS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   4935
      Begin VB.CommandButton cmdData 
         Caption         =   "Data Update"
         Height          =   975
         Left            =   3720
         Picture         =   "FrmMenu.frx":173A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdTaxReport 
         Caption         =   "Tax && Income Reports"
         Height          =   975
         Left            =   2520
         Picture         =   "FrmMenu.frx":1B7C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdTransfer 
         Caption         =   "Money Transfer"
         Height          =   975
         Left            =   1320
         Picture         =   "FrmMenu.frx":1FBE
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdForms 
         Caption         =   "Forms"
         Height          =   975
         Left            =   120
         Picture         =   "FrmMenu.frx":2400
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdInvoice 
         Caption         =   "Invoices"
         Height          =   975
         Left            =   2520
         Picture         =   "FrmMenu.frx":2842
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdPeople 
         Caption         =   "People / Entities"
         Height          =   975
         Left            =   120
         Picture         =   "FrmMenu.frx":2C84
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdWork 
         Caption         =   "Work Orders"
         Height          =   975
         Left            =   1320
         Picture         =   "FrmMenu.frx":30C6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit to Windows"
         Height          =   975
         Left            =   3720
         Picture         =   "FrmMenu.frx":3508
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1320
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdData_Click()
    Dim RS As ADODB.Recordset
    Set RS = RSTemplate("", "", 0, "Exec EastTreeUpdater")
    
    CloseRS RS
    Set RS = Nothing

End Sub

Private Sub cmdExit_Click()
Ending
End Sub

Private Sub cmdInvoice_Click()
LoadForm frmInvoice
End Sub

Private Sub cmdTaxReport_Click()
LoadForm frmTaxReport
End Sub

Private Sub cmdTransfer_Click()
LoadForm frmTransfer
End Sub

Private Sub cmdPeople_Click()
LoadForm frmPeople
End Sub

Private Sub cmdForms_Click()
LoadForm frmForms
End Sub

Private Sub cmdWork_Click()
LoadForm frmWorkOrder
End Sub

Private Sub Form_Activate()
UserClick = True
End Sub
Private Sub Form_Load()
lblversion.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Timer_Timer()
If Not Left(txtTime.Text, 2) = Left(Time$, 2) Then
    txtDate.Text = Date$
End If
txtTime.Text = Time$
End Sub
