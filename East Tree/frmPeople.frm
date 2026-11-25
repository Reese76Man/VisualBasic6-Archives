VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPeople 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "People & Entities"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11250
   Icon            =   "frmPeople.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   482
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   750
   Begin TabDlg.SSTab SSTab 
      Height          =   3975
      Left            =   0
      TabIndex        =   31
      Top             =   3240
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   7011
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Main Details"
      TabPicture(0)   =   "frmPeople.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameAddress"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameComments"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Work Orders"
      TabPicture(1)   =   "frmPeople.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "OPTWOAll"
      Tab(1).Control(2)=   "OPTWOOut"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Transactions"
      TabPicture(2)   =   "frmPeople.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   59
         Top             =   720
         Width           =   10935
         Begin VB.CommandButton cmdTransDelete 
            Caption         =   "Delete Transaction"
            Enabled         =   0   'False
            Height          =   255
            Left            =   8520
            TabIndex        =   67
            Top             =   2760
            Width           =   2295
         End
         Begin VB.TextBox txtStatusTrans 
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
            Left            =   3720
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   61
            Text            =   "No Transactions Found"
            Top             =   1320
            Width           =   3255
         End
         Begin VB.CommandButton cmdTransReport 
            Caption         =   "View / Print Report"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   2760
            Width           =   8295
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSDGTrans 
            Height          =   2610
            Left            =   120
            TabIndex        =   62
            Top             =   120
            Visible         =   0   'False
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   4604
            _Version        =   393216
            FixedCols       =   0
            ScrollTrack     =   -1  'True
            TextStyleFixed  =   1
            FocusRect       =   0
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
      Begin VB.OptionButton OPTWOOut 
         Caption         =   "Outstanding (Open Only)"
         Height          =   255
         Left            =   -73560
         TabIndex        =   58
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton OPTWOAll 
         Caption         =   "All"
         Height          =   255
         Left            =   -74760
         TabIndex        =   57
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Frame FrameComments 
         Height          =   3375
         Left            =   4560
         TabIndex        =   52
         Top             =   360
         Width           =   6495
         Begin VB.TextBox txtRate 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
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
            Left            =   1440
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   54
            Top             =   3000
            Width           =   1335
         End
         Begin VB.TextBox txtComments 
            Height          =   2535
            Left            =   120
            MaxLength       =   1024
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Top             =   240
            Width           =   6255
         End
         Begin VB.Label LabelRate 
            Alignment       =   1  'Right Justify
            Caption         =   "Tax/Cut Rate:"
            Height          =   255
            Left            =   240
            TabIndex        =   55
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000003&
            X1              =   120
            X2              =   6360
            Y1              =   2880
            Y2              =   2880
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   45
         Top             =   720
         Width           =   10935
         Begin VB.CommandButton cmdWorkOrder 
            Caption         =   "Jump to Work Order"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   2760
            Width           =   10695
         End
         Begin VB.TextBox txtStatusWork 
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
            Left            =   3720
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   46
            Text            =   "No Work Orders Found"
            Top             =   1320
            Width           =   3255
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSDGWork 
            Height          =   2610
            Left            =   120
            TabIndex        =   47
            Top             =   120
            Visible         =   0   'False
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   4604
            _Version        =   393216
            FixedCols       =   0
            ScrollTrack     =   -1  'True
            TextStyleFixed  =   1
            FocusRect       =   0
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
      Begin VB.Frame FrameAddress 
         Caption         =   "Contact Information"
         Height          =   3375
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   4335
         Begin VB.TextBox txtEmail 
            Height          =   285
            Left            =   1080
            MaxLength       =   255
            TabIndex        =   26
            Top             =   2640
            Width           =   3135
         End
         Begin VB.TextBox txtWebSite 
            Height          =   285
            Left            =   1080
            MaxLength       =   255
            TabIndex        =   27
            Top             =   3000
            Width           =   3135
         End
         Begin VB.TextBox txtPhone1 
            Height          =   285
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   24
            Top             =   1560
            Width           =   3135
         End
         Begin VB.TextBox txtPhone2 
            Height          =   285
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   25
            Top             =   1920
            Width           =   3135
         End
         Begin VB.TextBox txtZipcode 
            Height          =   285
            Left            =   3360
            MaxLength       =   50
            TabIndex        =   23
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtState 
            Height          =   285
            Left            =   2760
            MaxLength       =   50
            TabIndex        =   22
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtCity 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   21
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox txtAddress2 
            Height          =   285
            Left            =   1080
            MaxLength       =   100
            TabIndex        =   20
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox txtAddress1 
            Height          =   285
            Left            =   1080
            MaxLength       =   100
            TabIndex        =   19
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label9 
            Caption         =   "Phone 1:"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Phone 2:"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Email:"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label Label8 
            Caption         =   "Website:"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label Label13 
            Caption         =   "City/St/Zip:"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Address 2:"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Address 1:"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   855
         End
      End
   End
   Begin VB.Frame FrameMain 
      Height          =   1935
      Left            =   0
      TabIndex        =   35
      Top             =   1200
      Width           =   11175
      Begin VB.TextBox txtNumber 
         Height          =   285
         Left            =   5880
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtPaid 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
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
         Left            =   9720
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   18
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
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
         Left            =   9720
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cboEType 
         Height          =   315
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   3135
      End
      Begin VB.CheckBox chkActive 
         Caption         =   "Currently Active"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtOwe 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
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
         Left            =   9720
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtCompany 
         Height          =   285
         Left            =   960
         MaxLength       =   100
         TabIndex        =   14
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtSuffix 
         Height          =   285
         Left            =   7440
         MaxLength       =   25
         TabIndex        =   13
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtLast 
         Height          =   285
         Left            =   5040
         MaxLength       =   100
         TabIndex        =   12
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtMiddle 
         Height          =   285
         Left            =   3120
         MaxLength       =   100
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtID 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtFirst 
         Height          =   285
         Left            =   960
         MaxLength       =   100
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Alt. Number:"
         Height          =   255
         Left            =   4920
         TabIndex        =   53
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblPaid 
         Alignment       =   1  'Right Justify
         Caption         =   "Paid:"
         Height          =   255
         Left            =   8400
         TabIndex        =   51
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblOwe 
         Alignment       =   1  'Right Justify
         Caption         =   "Owe:"
         Height          =   255
         Left            =   8400
         TabIndex        =   50
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Type:"
         Height          =   255
         Left            =   4560
         TabIndex        =   49
         Top             =   240
         Width           =   495
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         X1              =   8280
         X2              =   8280
         Y1              =   240
         Y2              =   1800
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         Caption         =   "Balance:"
         Height          =   255
         Left            =   8400
         TabIndex        =   48
         Top             =   240
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   8280
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label7 
         Caption         =   "Company:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Suffix:"
         Height          =   255
         Left            =   6960
         TabIndex        =   34
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Last:"
         Height          =   255
         Left            =   4560
         TabIndex        =   33
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Middle:"
         Height          =   255
         Left            =   2520
         TabIndex        =   32
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "ID Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "First:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   495
      End
   End
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.CommandButton cmdCustomer 
         Caption         =   "Clients"
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
         Left            =   5760
         Picture         =   "frmPeople.frx":0496
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdGovernment 
         Caption         =   "Gvrnmnt"
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
         Left            =   6840
         Picture         =   "frmPeople.frx":08D8
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdVendor 
         Caption         =   "Vendors"
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
         Left            =   7920
         Picture         =   "frmPeople.frx":0D1A
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdEmployee 
         Caption         =   "Members"
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
         Left            =   4680
         Picture         =   "frmPeople.frx":115C
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   240
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
         Left            =   9000
         Picture         =   "frmPeople.frx":159E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
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
         Picture         =   "frmPeople.frx":19E0
         Style           =   1  'Graphical
         TabIndex        =   1
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
         Left            =   1200
         Picture         =   "frmPeople.frx":2122
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
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
         Left            =   2280
         Picture         =   "frmPeople.frx":2564
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Left            =   10200
         Picture         =   "frmPeople.frx":29A6
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   3360
         Picture         =   "frmPeople.frx":2DE8
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.Line Line5 
         X1              =   4440
         X2              =   4440
         Y1              =   960
         Y2              =   240
      End
      Begin VB.Line Line4 
         X1              =   10080
         X2              =   10080
         Y1              =   960
         Y2              =   240
      End
   End
End
Attribute VB_Name = "frmPeople"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FormCurrentValue As Long, AddNew As Boolean, EditChange As Boolean

Private Sub cmdAdd_Click()
If UserClick = True Then SetObjects 1
End Sub
Private Sub cmdCancel_Click()
If UserClick = True Then EntityLoad
End Sub

Private Sub Form_Activate()
txtFirst.SetFocus
End Sub

Private Sub Form_Load()
cboEType.Tag = FillBox(cboEType, "EntityType", "Description", "EntityTypeID")
cboEType.ListIndex = 0

FormCurrentValue = FormPassedValue
FormPassedValue = 0

SSTab.Tab = 0

EntityLoad
End Sub

Private Sub EntityLoad()
If FormCurrentValue = 0 Then FormCurrentValue = Val(TransTemplate("Select Top 1 EntityID From Entity Order By EntityID Asc", "EntityID", "0"))

Dim RS As ADODB.Recordset
Set RS = RSTemplate("", "", 0, "Exec EntityMain " & Str(FormCurrentValue))

If Not RS.EOF Then
    txtID.Text = RS!EntityNumber
    txtFirst.Text = RS!FirstName
    txtMiddle.Text = Trim("" & RS!MiddleName)
    txtLast.Text = RS!LastName
    txtSuffix.Text = Trim("" & RS!Suffix)
    txtEmail.Text = Trim("" & RS!EmailAddress)
    txtWebSite.Text = Trim("" & RS!WebSite)
    txtCompany.Text = Trim("" & RS!Company)
    txtPhone1.Text = PhoneFormat(("" & RS!Phone1), 1)
    txtPhone2.Text = PhoneFormat(("" & RS!Phone2), 1)
    txtAddress1.Text = Trim("" & RS!Address1)
    txtAddress2.Text = Trim("" & RS!Address2)
    txtCity.Text = Trim("" & RS!City)
    txtState.Text = Trim("" & RS!State)
    txtZipcode.Text = Trim("" & RS!Zipcode)
    chkActive.Value = Trim("" & RS!Active)
    cboEType.Tag = FindItem(cboEType, Trim("" & RS!EntityTypeDesc))
    txtNumber.Text = Trim("" & RS!AltNumber)
    txtRate.Text = Trim("" & RS!TaxRate)
    txtComments.Text = Trim("" & RS!Comments)
    
    txtBalance.Text = Trim("" & RS!Balance)
    txtOwe.Text = Trim("" & RS!Owed)
    txtPaid.Text = Trim("" & RS!Paid)
    
    If txtRate.Text = "" Then txtRate.Text = "N/A"
    
    Select Case Val("" & RS!EntityTypeID)
    Case 1  ' Employees
        lblBalance.Caption = "Payroll:"
        lblOwe.Caption = "Earned:"
        lblPaid.Caption = "Paid Out:"
        'txtRate.Text = "N/A"
    Case 2  ' Customers
        lblBalance.Caption = "Oustanding:"
        lblOwe.Caption = "Billed:"
        lblPaid.Caption = "Paid:"
        'txtRate.Text = "N/A"
    Case 5  ' Bank
        lblBalance.Caption = "Balance:"
        lblOwe.Caption = "N/A:"
        lblPaid.Caption = "Balance:"
        'txtRate.Text = "N/A"
    Case 3, 4, 6 ' Tax & Business Cuts
        lblBalance.Caption = "Balance:"
        lblOwe.Caption = "Owed:"
        lblPaid.Caption = "Paid:"
    Case Else
        lblBalance.Caption = "Balance:"
        lblOwe.Caption = "Owe:"
        lblPaid.Caption = "Paid:"
        'txtRate.Text = "N/A"
    End Select
    
    'Select Case Val(txtBalance.Text)
    'Case Is > 0
    '    txtBalance.BackColor = RGB(255, 175, 175)
    'Case Is < 0
    '    txtBalance.BackColor = RGB(175, 255, 175)
    'Case Else
    '    txtBalance.BackColor = &H80000018
    'End Select
End If

FillWOGrid FormCurrentValue
FillTAGrid FormCurrentValue

CloseRS RS
Set RS = Nothing

SetObjects 0
End Sub

Private Sub cmdUpdate_Click()
If UserClick = True And EditChange = True Then

StartRoutine Me
        Dim RS As ADODB.Recordset
        Set RS = RSTemplate("Entity", "EntityID", FormCurrentValue)
    
        If AddNew = True Then RS.AddNew
                
        RS!EntityTypeID = Trim(cboEType.ItemData(cboEType.ListIndex))
        RS!FirstName = Trim(txtFirst.Text)
        RS!MiddleName = Trim(txtMiddle.Text)
        RS!LastName = Trim(txtLast.Text)
        RS!Suffix = Trim(txtSuffix.Text)
        RS!Phone1 = Trim(txtPhone1.Text)
        RS!Phone2 = Trim(txtPhone2.Text)
        RS!Suffix = Trim(txtSuffix.Text)
        RS!Address1 = Trim(txtAddress1.Text)
        RS!Address2 = Trim(txtAddress2.Text)
        RS!City = Trim(txtCity.Text)
        RS!State = Trim(txtState.Text)
        RS!Zipcode = Trim(txtZipcode.Text)
        RS!EmailAddress = Trim(txtEmail.Text)
        RS!WebSite = Trim(txtWebSite.Text)
        RS!AltNumber = Trim(txtNumber.Text)
        RS!Active = chkActive.Value
        RS!Company = Trim(txtCompany.Text)
        'RS!TaxRate = Trim(txtRate.Text)
        
        RS!Comments = Trim(txtComments.Text)
        RS.Update
        
        If AddNew = True Then FormCurrentValue = RS!EntityID
    
        CloseRS RS
        Set RS = Nothing
    EndRoutine Me
    
    EntityLoad
End If
End Sub


Private Sub SetObjects(ComVar As Byte)
Select Case ComVar
Case 0  ' Normal
    UnLockTabs SSTab, SSTab.Tab
    FrameMain.Enabled = True
    cmdSearch.Enabled = True
    cmdEmployee.Enabled = True
    cmdGovernment.Enabled = True
    cmdVendor.Enabled = True
    cmdCustomer.Enabled = True
    cmdAdd.Enabled = True
    cmdCancel.Enabled = False
    cmdUpdate.Enabled = False
    cmdExit.Enabled = True
    'cmdDelete.Enabled = True
    
    AddEditFalse
Case 1  ' Add
    LockTabs SSTab, 0
    ECheck
    AddNew = True
    ClearFrame Me, FrameMain
    ClearFrame Me, FrameAddress
    ClearFrame Me, FrameComments
    
    txtFirst.SetFocus
End Select
End Sub

Private Sub cmdSearch_Click()
If UserClick = True Then
    cmdSearch.Tag = PickBox(0)
    
    If Val(cmdSearch.Tag) > 0 Then FormCurrentValue = Val(cmdSearch.Tag)
    EntityLoad
End If
End Sub
Private Sub cmdEmployee_Click()
If UserClick = True Then
    cmdSearch.Tag = PickBox(0, "[MemOnly]")
    
    If Val(cmdSearch.Tag) > 0 Then FormCurrentValue = Val(cmdSearch.Tag)
    EntityLoad
End If
End Sub
Private Sub cmdCustomer_Click()
If UserClick = True Then
    cmdSearch.Tag = PickBox(0, "[CusOnly]")
    
    If Val(cmdSearch.Tag) > 0 Then FormCurrentValue = Val(cmdSearch.Tag)
    EntityLoad
End If
End Sub
Private Sub cmdGovernment_Click()
If UserClick = True Then
    cmdSearch.Tag = PickBox(0, "[GovOnly]")
    
    If Val(cmdSearch.Tag) > 0 Then FormCurrentValue = Val(cmdSearch.Tag)
    EntityLoad
End If
End Sub
Private Sub cmdVendor_Click()
If UserClick = True Then
    cmdSearch.Tag = PickBox(0, "[VenOnly]")
    
    If Val(cmdSearch.Tag) > 0 Then FormCurrentValue = Val(cmdSearch.Tag)
    EntityLoad
End If
End Sub

Private Sub cmdExit_Click()
If UserClick = True Then Unload Me
End Sub

' ************************************** Grids ************************************
Private Sub FillWOGrid(ComID As Long)
StartRoutine Me
    If OPTWOOut.Value = True Then
        OPTWOAll.Tag = 1    ' Outstanding (Open Only)
    Else
        OPTWOAll.Tag = 0    ' All
    End If

    If FillGrid("Exec WorkOrderGrid " & Str(ComID) & "," & Str(OPTWOAll.Tag), MSDGWork, "", txtStatusWork) = True Then
        With MSDGWork
            .ColWidth(1) = 750      ' OrderID
            .ColWidth(2) = 1000     ' Open Date
            .ColWidth(3) = 250      ' Relation
            .ColWidth(4) = 650      ' EmpID
            .ColWidth(5) = 1250     ' Employee
            .ColWidth(6) = 0        ' Emp Company
            .ColWidth(7) = 650      ' CusID
            .ColWidth(8) = 1350     ' Customer
            .ColWidth(9) = 2100     ' Cus Company
            .ColWidth(10) = 650     ' Charged
            .ColWidth(11) = 650     ' Paid
            .ColWidth(12) = 1000    ' Close Date
            .Row = 1
            .Col = 1
        End With

        cmdWorkOrder.Enabled = True
    Else
        cmdWorkOrder.Enabled = False
    End If

    MSDGWork_Click
EndRoutine Me
End Sub

Private Sub MSDGWork_Click()
GridFix MSDGWork: DoEvents
If MSDGWork.Row > 0 And MSDGWork.TextMatrix(MSDGWork.Row, 0) <> "" Then
    MSDGWork.Tag = Trim(MSDGWork.TextMatrix(MSDGWork.Row, 0))
End If
End Sub
Private Sub MSDGWork_DblClick()
MSDGWork_Click
cmdWorkOrder_Click
End Sub
Private Sub MSDGWork_RowColChange()
MSDGWork_Click
End Sub

Private Sub OPTWOAll_Click()
If UserClick = True Then FillWOGrid FormCurrentValue
End Sub
Private Sub OPTWOOut_Click()
OPTWOAll_Click
End Sub

Private Sub cmdWorkOrder_Click()
If UserClick = True Then LoadForm frmWorkOrder, , Val(MSDGWork.Tag)
End Sub

'**************

Private Sub FillTAGrid(ComID As Long)
StartRoutine Me
    'If OPTWOOut.Value = True Then
    '    OPTWOAll.Tag = 1    ' Outstanding (Open Only)
    'Else
    '    OPTWOAll.Tag = 0    ' All
    'End If

    If FillGrid("Exec TransactionGrid " & Str(ComID), MSDGTrans, "", txtStatusTrans) = True Then
        With MSDGTrans
            .ColWidth(1) = 750      ' OrderID
            .ColWidth(2) = 1000     ' Date
            .ColWidth(3) = 250      ' Relation
            .ColWidth(4) = 650      ' PayerID
            .ColWidth(5) = 1300     ' Employee
            .ColWidth(6) = 1600     ' Emp Company
            .ColWidth(7) = 650      ' PayeeID
            .ColWidth(8) = 1350     ' Customer
            .ColWidth(9) = 2100     ' Cus Company
            .ColWidth(10) = 650     ' Charged
            '.ColWidth(11) = 650     ' Paid
            '.ColWidth(12) = 1000    ' Close Date
            .Row = 1
            .Col = 1
        End With

        cmdTransReport.Enabled = True
        cmdTransDelete.Enabled = True
    Else
        cmdTransReport.Enabled = False
        cmdTransDelete.Enabled = False
    End If

    MSDGTrans_Click
EndRoutine Me
End Sub

Private Sub MSDGTrans_Click()
GridFix MSDGTrans: DoEvents
If MSDGTrans.Row > 0 And MSDGTrans.TextMatrix(MSDGTrans.Row, 0) <> "" Then
    MSDGTrans.Tag = Trim(MSDGTrans.TextMatrix(MSDGTrans.Row, 0))
End If
End Sub
Private Sub MSDGTrans_DblClick()
MSDGTrans_Click
cmdTransReport_Click
End Sub
Private Sub MSDGTrans_RowColChange()
MSDGTrans_Click
End Sub

Private Sub cmdTransReport_Click()
If UserClick = True Then
    StartRoutine Me
        Dim RS As ADODB.Recordset
        Set RS = RSTemplate("", "", 0, "Exec TransactionGrid " & Str(FormCurrentValue))
        
        frmReport.ViewReport "TransactionReport", RS, , False
        LoadForm frmReport, 1
        
        CloseRS RS
        Set RS = Nothing
    EndRoutine Me
End If
End Sub
Private Sub cmdTransDelete_Click()
If UserClick = True And Val(MSDGTrans.TextMatrix(MSDGTrans.Row, 1)) = 0 And Val(MSDGTrans.Tag) > 0 Then
    If MsgBox("Are you sure that you want to delete this transaction?", vbYesNo, "Fund Transfer") = vbYes Then
        DeleteTemplate "WorkDetail", "WorkDetailID", Val(MSDGTrans.Tag)
        
        MsgBox "Transaction Deleted!", vbInformation
        
        Dim RS As ADODB.Recordset
        Set RS = RSTemplate("", "", 0, "Exec EastTreeUpdater 0")
        Set RS = Nothing
        
        FillTAGrid FormCurrentValue
    End If
End If
End Sub
' *************************************************************************************************

Private Sub ECheck()
EditChange = EditCheck(cmdUpdate, cmdAdd, cmdCancel, cmdDelete, cmdExit, cmdSearch, cmdEmployee, cmdGovernment, cmdVendor, cmdCustomer)
End Sub
Private Sub chkActive_Click()
ECheck
End Sub
Private Sub txtAddress1_Change()
ECheck
End Sub
Private Sub txtAddress2_Change()
ECheck
End Sub
Private Sub txtCity_Change()
ECheck
End Sub
Private Sub txtCompany_Change()
ECheck
End Sub
Private Sub txtEmail_Change()
ECheck
End Sub
Private Sub txtFirst_Change()
ECheck
End Sub
Private Sub txtLast_Change()
ECheck
End Sub
Private Sub txtMiddle_Change()
ECheck
End Sub
Private Sub txtPhone1_Change()
ECheck
End Sub
Private Sub txtPhone2_Change()
ECheck
End Sub
Private Sub txtState_Change()
ECheck
End Sub
Private Sub txtSuffix_Change()
ECheck
End Sub
Private Sub cboEType_Click()
ECheck
End Sub
Private Sub txtWebSite_Change()
ECheck
End Sub
Private Sub txtZipcode_Change()
ECheck
End Sub
Private Sub txtNumber_Change()
ECheck
End Sub
Private Sub txtComments_Change()
ECheck
End Sub

Public Sub AddEditFalse()
    AddNew = False
    EditChange = False
End Sub
