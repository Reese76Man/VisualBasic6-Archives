VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmGroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Groups"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   Icon            =   "frmGroup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   6765
   Begin VB.Frame FrameGroups 
      Caption         =   "Group Memberships"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6585
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6735
      Begin VB.Frame FrameButtons 
         Caption         =   "Menu Options:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6495
         Begin VB.CommandButton cmdAdd 
            Caption         =   "New"
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
            Left            =   120
            Picture         =   "frmGroup.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "Save"
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
            Left            =   1920
            Picture         =   "frmGroup.frx":0B84
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete"
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
            Left            =   3120
            Picture         =   "frmGroup.frx":0FC6
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   855
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
            Height          =   735
            Left            =   5520
            Picture         =   "frmGroup.frx":1408
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
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
            Left            =   4320
            Picture         =   "frmGroup.frx":184A
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   855
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000003&
            X1              =   1440
            X2              =   1440
            Y1              =   960
            Y2              =   240
         End
      End
      Begin VB.Frame FrameGroupName 
         Caption         =   "Group Description"
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   6495
         Begin VB.TextBox txtGroup 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   6255
         End
      End
      Begin VB.TextBox txtGroupMembershipStatus 
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
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "No Groups Found"
         Top             =   4200
         Width           =   3855
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSDGGroups 
         Height          =   4260
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   7514
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   8388608
         ScrollTrack     =   -1  'True
         TextStyleFixed  =   1
         FocusRect       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "frmGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
If UserClick = True Then SetObjects 1
End Sub

Private Sub SetObjects(ComVar As Byte)
If ComVar = 0 Then              ' Data Displayed (Reset to Normal View)
    cmdAdd.Enabled = True
    cmdExit.Enabled = True
    cmdCancel.Enabled = False
    AddNew = False
    EditChange = False
ElseIf ComVar = 1 Then          ' Add Engaged
    cmdAdd.Enabled = False
    cmdCancel.Enabled = True
    cmdDelete.Enabled = False
    cmdExit.Enabled = False
    EditCheck cmdUpdate
    AddNew = True
    txtGroup.Text = ""
    txtGroup.SetFocus
End If
End Sub

Private Sub cmdDelete_Click()
If UserClick = True Then
    MsgBox "This group may be associated with multiple entries!", vbExclamation
    If MsgBox("Are you sure you want to delete this Group?", vbYesNo) = vbYes Then
        DeleteTemplate "MembershipGroupAssociation", "MembershipGroupID", MSDGGroups.Tag
        DeleteTemplate "MembershipGroup", "MembershipGroupID", MSDGGroups.Tag
        GroupUpdate
    End If
End If
End Sub

Private Sub cmdCancel_Click()
If UserClick = True Then GroupUpdate
End Sub

Private Sub cmdUpdate_Click()
If UserClick = True And EditChange = True And Len(Trim(txtGroup.Text)) > 0 Then
    StartRoutine Me
    Dim RS As Recordset
    Set RS = SaveTemplate("MembershipGroup", "MembershipGroupID", MSDGGroups.Tag)
    
    If AddNew = True Then RS.AddNew
    
    RS!Description = Trim(txtGroup.Text)
    RS.Update
    
    CloseRS RS
    Set RS = Nothing
    EndRoutine Me
    
    GroupUpdate
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Xout ""
End Sub

Private Sub cmdExit_Click()
If UserClick = True Then CloseWindow
End Sub

Private Sub Form_Load()
GroupUpdate
End Sub

Private Sub GridSet(GridObj As MSHFlexGrid, Optional GridNumber As Byte)
With GridObj
  If GridNumber = 0 Then
    .FormatString = "<|<|"
    .Cols = 2
    .ColWidth(0) = 0
    .ColWidth(1) = 7000
    .TextMatrix(0, 1) = "Group Membership"
  End If
  .Row = 1
  .Col = 1
End With
End Sub

Private Sub GroupUpdate()
DE.MembershipGroup
    If Not DE.rsMembershipGroup.EOF Then
        Set MSDGGroups.DataSource = DE.rsMembershipGroup
        txtGroupMembershipStatus.Visible = False
        MSDGGroups.Visible = True
        cmdDelete.Enabled = True
    Else
        
        cmdDelete.Enabled = False
        txtGroupMembershipStatus.Text = "No Groups Found"
        txtGroupMembershipStatus.Visible = True
        MSDGGroups.Visible = False
    End If
DE.rsMembershipGroup.Close

GridSet MSDGGroups, 0

SetObjects 0

MSDGGroups_Click
End Sub

Private Sub MSDGGroups_Click()
StartRoutine Me
    GridFix MSDGGroups
If MSDGGroups.Row > 0 And MSDGGroups.TextMatrix(MSDGGroups.Row, 0) <> "" Then
    MSDGGroups.Tag = Trim(MSDGGroups.TextMatrix(MSDGGroups.Row, 0))
    txtGroup.Text = Trim(MSDGGroups.TextMatrix(MSDGGroups.Row, 1))
End If
EndRoutine Me
End Sub
Private Sub MSDGGroups_DblClick()
MSDGGroups_Click
End Sub
Private Sub MSDGGroups_RowColChange()
MSDGGroups_Click
End Sub

Private Sub txtgroup_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
End Sub
