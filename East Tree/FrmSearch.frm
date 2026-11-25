VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Box"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameSearch 
      Caption         =   "Results:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   7695
      Begin VB.TextBox txtSearchStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "LastName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "No Search Results Found"
         Top             =   1680
         Width           =   3135
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSDGSearch 
         Height          =   3090
         Left            =   45
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   5450
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         TextStyleFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search Criteria (Must use at least 3 characters in criteria)"
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4815
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   320
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      Picture         =   "FrmSearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click to cancel an Add or Data Change."
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      Picture         =   "FrmSearch.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Picture         =   "FrmSearch.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Click to save data."
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FrmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PickType As Long

Private Sub InitSearch(ComNum As String)
StartRoutine Me

    If FillGrid("Exec Search " & PickType & ", '" & ComNum & "'", MSDGSearch, "Search Results", txtSearchStatus) = True Then
        With MSDGSearch   '7200 Total
        
            Select Case PickType
            Case 7      ' Work Order
                .ColWidth(1) = 3500
                .ColWidth(2) = 800
                .ColWidth(3) = 800
                .ColWidth(4) = 2100
            Case Else   ' People, etc.
                .ColWidth(1) = 2600
                .ColWidth(2) = 1500
                .ColWidth(3) = 2500
                .ColWidth(4) = 600
            End Select
            .Row = 1
            .Col = 1
        End With

        
        GridFix MSDGSearch
    End If
    
    cmdSelect.Enabled = False
EndRoutine Me
End Sub

Private Sub cmdCancel_Click()
FormPassedValue = 0
Unload Me
End Sub

Private Sub cmdSelect_Click()
If UserClick = True Then
    FormPassedValue = Trim(MSDGSearch.TextMatrix(MSDGSearch.Row, 0))
    Unload Me
End If
End Sub

Private Sub cmdSearch_Click()
If UserClick = True And Len(txtSearch.Text) > 2 Then
    InitSearch SQLFix(txtSearch.Text)
    MSDGSearch_Click
End If
End Sub

Private Sub Form_Activate()
txtSearch.SetFocus
End Sub

Private Sub Form_Load()
UserClick = True
PickType = FormPassedValue: DoEvents
If Len(FormPassedText) > 0 Then
    txtSearch.Text = FormPassedText: DoEvents
    InitSearch SQLFix(txtSearch.Text)
End If
End Sub

Private Sub MSDGSearch_Click()
GridFix MSDGSearch: DoEvents
If UserClick = True And MSDGSearch.Row > 0 And MSDGSearch.TextMatrix(MSDGSearch.Row, 0) <> "" Then
    cmdSelect.Enabled = True
    MSDGSearch.SetFocus
End If
End Sub

Private Sub MSDGSearch_DblClick()
MSDGSearch_Click
DoEvents
cmdSelect_Click
End Sub

Private Sub MSDGSearch_RowColChange()
MSDGSearch_Click
End Sub

Private Sub txtSearch_Change()
If Len(txtSearch.Text) > 2 Then
    cmdSearch.Enabled = True
Else
    cmdSearch.Enabled = False
End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And cmdSearch.Enabled = True Then
    KeyAscii = 0
    cmdSearch_Click
    Exit Sub
End If
End Sub
Private Sub MSDGSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And cmdSelect.Enabled = True Then
    KeyAscii = 0
    cmdSelect_Click
    Exit Sub
End If
End Sub

