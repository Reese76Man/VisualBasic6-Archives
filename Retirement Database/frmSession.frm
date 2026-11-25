VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retirees"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11400
   Icon            =   "frmSession.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   496
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   760
   Begin VB.Frame FrameCourseHeader 
      Caption         =   "Retiree Primary Information:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   66
      Top             =   1200
      Width           =   11415
      Begin VB.Frame FrameChapters 
         Height          =   2055
         Left            =   6840
         TabIndex        =   102
         Top             =   120
         Width           =   4455
         Begin VB.TextBox txtNationalChapterNumber 
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
            Left            =   3120
            MaxLength       =   50
            TabIndex        =   110
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtChapterStatus 
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
            Left            =   360
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   106
            Text            =   "No Local Chapters Found"
            Top             =   1440
            Width           =   3015
         End
         Begin VB.CommandButton cmdAddChapter 
            Caption         =   "Add"
            Height          =   310
            Left            =   3720
            TabIndex        =   105
            Top             =   1200
            Width           =   615
         End
         Begin VB.CommandButton cmdDeleteChapter 
            Caption         =   "Delete"
            Height          =   310
            Left            =   3720
            TabIndex        =   104
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txtLocalChapterNumber 
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
            Left            =   3120
            MaxLength       =   50
            TabIndex        =   103
            Top             =   840
            Width           =   1215
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSDGChapters 
            Height          =   780
            Left            =   120
            TabIndex        =   107
            Top             =   1200
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   1376
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
         Begin MSComCtl2.DTPicker DTPLocal 
            Height          =   315
            Left            =   1440
            TabIndex        =   108
            Top             =   840
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   20381697
            CurrentDate     =   38721
         End
         Begin MSComCtl2.DTPicker DTPNational 
            Height          =   315
            Left            =   1440
            TabIndex        =   111
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   20381697
            CurrentDate     =   38721
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000003&
            BorderWidth     =   2
            X1              =   0
            X2              =   4430
            Y1              =   750
            Y2              =   750
         End
         Begin VB.Label lblOtherState2 
            Caption         =   "National Chapter:"
            Height          =   255
            Left            =   120
            TabIndex        =   114
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Joined Date"
            Height          =   255
            Left            =   1440
            TabIndex        =   113
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Chapter"
            Height          =   255
            Left            =   3120
            TabIndex        =   112
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label lblOtherState1 
            Caption         =   "Local Chapter(s):"
            Height          =   255
            Left            =   120
            TabIndex        =   109
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.CheckBox chkNonMember 
         Caption         =   "Non-Member"
         Height          =   255
         Left            =   2640
         TabIndex        =   99
         Top             =   1950
         Width           =   1215
      End
      Begin VB.CheckBox chkDeceased 
         Caption         =   "Deceased"
         Height          =   255
         Left            =   5640
         TabIndex        =   97
         Top             =   1950
         Width           =   1095
      End
      Begin VB.CheckBox chkWidowed 
         Caption         =   "Widow"
         Height          =   255
         Left            =   4320
         TabIndex        =   96
         Top             =   1950
         Width           =   975
      End
      Begin VB.CheckBox chkActive 
         Caption         =   "Active"
         Height          =   255
         Left            =   1440
         TabIndex        =   95
         Top             =   1950
         Width           =   855
      End
      Begin VB.CheckBox chkBus 
         Caption         =   "Business Address"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   85
         Top             =   880
         Width           =   1095
      End
      Begin VB.CommandButton cmdSendPEmail 
         Caption         =   "Send"
         Height          =   285
         Left            =   6120
         TabIndex        =   24
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtEmail 
         BackColor       =   &H00FFFFFF&
         DataField       =   "MiddleName"
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
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   23
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtPhone 
         BackColor       =   &H00FFFFFF&
         DataField       =   "MiddleName"
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
         MaxLength       =   50
         TabIndex        =   22
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtCountry 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   21
         Top             =   1275
         Width           =   1335
      End
      Begin VB.TextBox txtAddress1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "MiddleName"
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
         MaxLength       =   150
         TabIndex        =   16
         Top             =   600
         Width           =   5295
      End
      Begin VB.TextBox txtAddress2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "MiddleName"
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
         MaxLength       =   150
         TabIndex        =   17
         Top             =   940
         Width           =   3975
      End
      Begin VB.TextBox txtState 
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
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   19
         Top             =   1275
         Width           =   495
      End
      Begin VB.TextBox txtCity 
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
         MaxLength       =   75
         TabIndex        =   18
         Top             =   1275
         Width           =   2055
      End
      Begin VB.TextBox txtZipcode 
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
         Left            =   3960
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   20
         Top             =   1275
         Width           =   1455
      End
      Begin VB.TextBox txtFirstName 
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
         MaxLength       =   50
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtPrefix 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         MaxLength       =   10
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtSuffix 
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
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtMiddleName 
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
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtLastName 
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
         Left            =   4800
         MaxLength       =   50
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Member Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   1950
         Width           =   1215
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   6840
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Email:"
         Height          =   255
         Left            =   3120
         TabIndex        =   71
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Phone:"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "City/ST/Zip/Ctry:"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   1275
         Width           =   1335
      End
      Begin VB.Label lblPrimaryAddress1 
         Caption         =   "Primary Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   600
         Width           =   1455
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000003&
         X1              =   6840
         X2              =   0
         Y1              =   1900
         Y2              =   1900
      End
      Begin VB.Label Label4 
         Caption         =   "Retiree Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame FrameSearch 
      Caption         =   "Search Options:"
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
      Left            =   5160
      TabIndex        =   61
      Top             =   0
      Width           =   6255
      Begin VB.ComboBox cboSearch 
         Height          =   315
         ItemData        =   "frmSession.frx":0442
         Left            =   120
         List            =   "frmSession.frx":0449
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   9
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton cmdRetrieve 
         Caption         =   "Retrieve"
         Height          =   285
         Left            =   5280
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   285
         Left            =   5280
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin MSDataListLib.DataCombo cboEntry 
         Bindings        =   "frmSession.frx":045E
         Height          =   315
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "FullName"
         BoundColumn     =   "FullName"
         Text            =   ""
         Object.DataMember      =   "EntryFullNames"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   6120
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label11 
         Caption         =   "Retiree:"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame FrameButtons 
      Caption         =   "Retiree Menu Options:"
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
      Width           =   5055
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
         Left            =   3120
         Picture         =   "frmSession.frx":0470
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   4080
         Picture         =   "frmSession.frx":08B2
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   2160
         Picture         =   "frmSession.frx":0CF4
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmSession.frx":1136
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmSession.frx":1578
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         X1              =   1080
         X2              =   1080
         Y1              =   960
         Y2              =   240
      End
   End
   Begin TabDlg.SSTab SSTAB 
      Height          =   3855
      Left            =   0
      TabIndex        =   25
      Top             =   3600
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Retiree Additional Information"
      TabPicture(0)   =   "frmSession.frx":1CBA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ShapeBox"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "EntryImage"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameCourseInformation"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdPicture"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Retiree Addresses && Emails"
      TabPicture(1)   =   "frmSession.frx":1CD6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Retiree Dues"
      TabPicture(2)   =   "frmSession.frx":1CF2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Comments"
      TabPicture(3)   =   "frmSession.frx":1D0E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrameComments"
      Tab(3).ControlCount=   1
      Begin VB.Frame FrameComments 
         Caption         =   "Comments"
         Height          =   3450
         Left            =   -74880
         TabIndex        =   100
         Top             =   360
         Width           =   11175
         Begin VB.TextBox txtComments 
            Height          =   3135
            Left            =   120
            MaxLength       =   8192
            MultiLine       =   -1  'True
            TabIndex        =   101
            Top             =   240
            Width           =   10935
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3450
         Left            =   -74880
         TabIndex        =   55
         Top             =   380
         Width           =   11175
         Begin VB.Frame FrameDues 
            Caption         =   "Dues"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2625
            Left            =   120
            TabIndex        =   76
            Top             =   720
            Width           =   10935
            Begin VB.CommandButton cmdUpdateDues 
               Caption         =   "Save"
               Height          =   555
               Left            =   10080
               TabIndex        =   87
               Top             =   1080
               Width           =   735
            End
            Begin VB.CommandButton cmdDelDues 
               Caption         =   "Delete"
               Height          =   675
               Left            =   10080
               TabIndex        =   53
               Top             =   1800
               Width           =   735
            End
            Begin VB.CommandButton cmdAddDues 
               Caption         =   "New"
               Height          =   675
               Left            =   10080
               TabIndex        =   52
               Top             =   240
               Width           =   735
            End
            Begin VB.TextBox txtDuesStatus 
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
               Left            =   3480
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   77
               Text            =   "No DuesFound"
               Top             =   1320
               Width           =   3855
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSDGDues 
               Height          =   2265
               Left            =   45
               TabIndex        =   51
               Top             =   240
               Width           =   9975
               _ExtentX        =   17595
               _ExtentY        =   3995
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
         Begin VB.Frame FrameDuesText 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   120
            TabIndex        =   88
            Top             =   240
            Width           =   10935
            Begin VB.TextBox txtDues 
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
               Left            =   1080
               MaxLength       =   50
               TabIndex        =   90
               Top             =   120
               Width           =   1215
            End
            Begin VB.TextBox txtDuesComments 
               BackColor       =   &H00FFFFFF&
               DataField       =   "MiddleName"
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
               Left            =   6360
               MaxLength       =   150
               TabIndex        =   89
               Top             =   120
               Width           =   4575
            End
            Begin MSComCtl2.DTPicker DTPDues 
               Height          =   315
               Left            =   3480
               TabIndex        =   91
               Top             =   120
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               Format          =   20381697
               CurrentDate     =   38721
            End
            Begin VB.Label Label14 
               Caption         =   "Dues Paid:  $"
               Height          =   255
               Left            =   120
               TabIndex        =   94
               Top             =   120
               Width           =   1095
            End
            Begin VB.Label Label15 
               Caption         =   "Date Paid:"
               Height          =   255
               Left            =   2640
               TabIndex        =   93
               Top             =   120
               Width           =   855
            End
            Begin VB.Label Label16 
               Caption         =   "Comments:"
               Height          =   255
               Left            =   5400
               TabIndex        =   92
               Top             =   120
               Width           =   975
            End
         End
      End
      Begin VB.CommandButton cmdPicture 
         Caption         =   "Add / Change Picture"
         Height          =   310
         Left            =   8160
         TabIndex        =   86
         Top             =   3500
         Width           =   3135
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3450
         Left            =   -74880
         TabIndex        =   65
         Top             =   380
         Width           =   11175
         Begin VB.Frame FrameEmailAddress 
            Caption         =   "Email Addresses"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2835
            Left            =   6960
            TabIndex        =   74
            Top             =   540
            Width           =   4095
            Begin VB.CommandButton cmdUpdateEmail 
               Caption         =   "Save"
               Height          =   375
               Left            =   3240
               TabIndex        =   47
               Top             =   720
               Width           =   735
            End
            Begin VB.CommandButton cmdSetPrimaryEmail 
               Caption         =   "Set As Primary"
               Height          =   495
               Left            =   3240
               TabIndex        =   49
               Top             =   1800
               Width           =   735
            End
            Begin VB.CommandButton cmdSendEmail 
               Caption         =   "Send Email"
               Height          =   495
               Left            =   3240
               TabIndex        =   48
               Top             =   1200
               Width           =   735
            End
            Begin VB.CommandButton cmdDelEmail 
               Caption         =   "Delete"
               Height          =   375
               Left            =   3240
               TabIndex        =   50
               Top             =   2400
               Width           =   735
            End
            Begin VB.CommandButton cmdAddEmail 
               Caption         =   "New"
               Height          =   375
               Left            =   3240
               TabIndex        =   46
               Top             =   240
               Width           =   735
            End
            Begin VB.TextBox txtEmailStatus 
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
               Left            =   960
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   75
               Text            =   "No Emails Found"
               Top             =   1200
               Width           =   1695
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSDGEmailAddress 
               Height          =   2505
               Left            =   120
               TabIndex        =   45
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   4419
               _Version        =   393216
               BackColor       =   -2147483639
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
         Begin VB.Frame FrameAddress 
            Caption         =   "Addresses"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1875
            Left            =   120
            TabIndex        =   72
            Top             =   1500
            Width           =   6615
            Begin VB.CommandButton cmdUpdateAddress 
               Caption         =   "Save"
               Height          =   310
               Left            =   5760
               TabIndex        =   41
               Top             =   600
               Width           =   735
            End
            Begin VB.CommandButton cmdSetPrimaryAddress 
               Caption         =   "Set As Primary"
               Height          =   425
               Left            =   5760
               TabIndex        =   42
               Top             =   960
               Width           =   735
            End
            Begin VB.TextBox txtAddressStatus 
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
               Left            =   1080
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   73
               Text            =   "No Addresses Found"
               Top             =   960
               Width           =   3975
            End
            Begin VB.CommandButton cmdAddAddress 
               Caption         =   "New"
               Height          =   310
               Left            =   5760
               TabIndex        =   40
               Top             =   240
               Width           =   735
            End
            Begin VB.CommandButton cmdDelAddress 
               Caption         =   "Delete"
               Height          =   310
               Left            =   5760
               TabIndex        =   43
               Top             =   1440
               Width           =   735
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSDGAddress 
               Height          =   1545
               Left            =   120
               TabIndex        =   39
               Top             =   240
               Width           =   5535
               _ExtentX        =   9763
               _ExtentY        =   2725
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
         Begin VB.Frame FrameEmailText 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   6960
            TabIndex        =   83
            Top             =   120
            Width           =   4095
            Begin VB.TextBox txtEEMail 
               BackColor       =   &H00FFFFFF&
               DataField       =   "MiddleName"
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
               Left            =   600
               MaxLength       =   200
               TabIndex        =   44
               Top             =   120
               Width           =   3495
            End
            Begin VB.Label lblEEmail 
               Caption         =   "Email:"
               Height          =   255
               Left            =   120
               TabIndex        =   84
               Top             =   120
               Width           =   615
            End
         End
         Begin VB.Frame FrameAddressText 
            BorderStyle     =   0  'None
            Height          =   1455
            Left            =   120
            TabIndex        =   78
            Top             =   120
            Width           =   6615
            Begin VB.OptionButton OptBusNo 
               Caption         =   "No"
               Height          =   220
               Left            =   5880
               TabIndex        =   38
               Top             =   1110
               Width           =   615
            End
            Begin VB.TextBox txtSZip 
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
               MaxLength       =   20
               TabIndex        =   34
               Top             =   750
               Width           =   1455
            End
            Begin VB.TextBox txtSCity 
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
               MaxLength       =   75
               TabIndex        =   32
               Top             =   750
               Width           =   1935
            End
            Begin VB.TextBox txtSState 
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
               Left            =   3240
               MaxLength       =   2
               TabIndex        =   33
               Top             =   750
               Width           =   495
            End
            Begin VB.TextBox txtSAddress2 
               BackColor       =   &H00FFFFFF&
               DataField       =   "MiddleName"
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
               MaxLength       =   150
               TabIndex        =   31
               Top             =   435
               Width           =   5295
            End
            Begin VB.TextBox txtSAddress1 
               BackColor       =   &H00FFFFFF&
               DataField       =   "MiddleName"
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
               MaxLength       =   150
               TabIndex        =   30
               Top             =   120
               Width           =   5295
            End
            Begin VB.TextBox txtSCountry 
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
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
               Left            =   5280
               MaxLength       =   50
               TabIndex        =   35
               Top             =   750
               Width           =   1335
            End
            Begin VB.TextBox txtSPhone 
               BackColor       =   &H00FFFFFF&
               DataField       =   "MiddleName"
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
               MaxLength       =   50
               TabIndex        =   36
               Top             =   1050
               Width           =   2295
            End
            Begin VB.OptionButton OptBusYes 
               Caption         =   "Yes"
               Height          =   220
               Left            =   5280
               TabIndex        =   37
               Top             =   1110
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.Shape Shape3 
               BorderColor     =   &H8000000C&
               Height          =   270
               Left            =   3720
               Top             =   1080
               Width           =   2895
            End
            Begin VB.Label Label12 
               Caption         =   "Address:"
               Height          =   255
               Left            =   0
               TabIndex        =   82
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label10 
               Caption         =   "City/ST/Zip/Ctry:"
               Height          =   255
               Left            =   0
               TabIndex        =   81
               Top             =   750
               Width           =   1455
            End
            Begin VB.Label Label9 
               Caption         =   "Phone:"
               Height          =   255
               Left            =   0
               TabIndex        =   80
               Top             =   1050
               Width           =   1335
            End
            Begin VB.Label Label13 
               Caption         =   "Business Address:"
               Height          =   255
               Left            =   3840
               TabIndex        =   79
               Top             =   1110
               Width           =   1335
            End
         End
         Begin VB.Line Line6 
            BorderColor     =   &H80000003&
            X1              =   6840
            X2              =   6840
            Y1              =   105
            Y2              =   3435
         End
      End
      Begin VB.Frame FrameCourseInformation 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3450
         Left            =   120
         TabIndex        =   54
         Top             =   380
         Width           =   7935
         Begin VB.Frame FrameOccupations 
            Caption         =   "Occupations && Classifications"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   120
            TabIndex        =   115
            Top             =   1920
            Width           =   7695
            Begin VB.TextBox txtOccupationStatus 
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
               Left            =   2280
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   121
               Text            =   "No Local Chapters Found"
               Top             =   840
               Width           =   3015
            End
            Begin VB.CommandButton cmdAddOccupation 
               Caption         =   "Add"
               Height          =   435
               Left            =   6960
               TabIndex        =   120
               Top             =   600
               Width           =   615
            End
            Begin VB.CommandButton cmdDeleteOccupation 
               Caption         =   "Delete"
               Height          =   310
               Left            =   6960
               TabIndex        =   119
               Top             =   1080
               Width           =   615
            End
            Begin VB.TextBox txtOccupationDesc 
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
               Left            =   2040
               MaxLength       =   50
               TabIndex        =   117
               Top             =   240
               Width           =   3015
            End
            Begin VB.TextBox txtOccupationClass 
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
               Left            =   5160
               MaxLength       =   50
               TabIndex        =   116
               Top             =   240
               Width           =   2415
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSDGOccupations 
               Height          =   780
               Left            =   120
               TabIndex        =   122
               Top             =   600
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   1376
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
            Begin VB.Label Label5 
               Caption         =   "Occupation Desc./Class:"
               Height          =   255
               Left            =   120
               TabIndex        =   118
               Top             =   240
               Width           =   2895
            End
         End
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
            Height          =   1785
            Left            =   120
            TabIndex        =   63
            Top             =   120
            Width           =   7695
            Begin VB.CommandButton cmdDeleteGroup 
               Caption         =   "Delete"
               Height          =   310
               Left            =   6960
               TabIndex        =   28
               Top             =   240
               Width           =   615
            End
            Begin VB.CommandButton cmdAddGroup 
               Caption         =   "Add"
               Height          =   310
               Left            =   6240
               TabIndex        =   27
               Top             =   240
               Width           =   615
            End
            Begin VB.ComboBox cboGroup 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmSession.frx":1D2A
               Left            =   120
               List            =   "frmSession.frx":1D2C
               Style           =   2  'Dropdown List
               TabIndex        =   26
               Top             =   240
               Width           =   6015
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
               Left            =   2040
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   64
               Text            =   "No Groups Found"
               Top             =   960
               Width           =   3855
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSDGGroups 
               Height          =   1140
               Left            =   45
               TabIndex        =   29
               Top             =   600
               Width           =   7575
               _ExtentX        =   13361
               _ExtentY        =   2011
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
      Begin VB.Image EntryImage 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3135
         Left            =   8160
         Stretch         =   -1  'True
         Top             =   380
         Width           =   3135
      End
      Begin VB.Shape ShapeBox 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         Height          =   3135
         Left            =   8160
         Top             =   380
         Width           =   3135
      End
   End
   Begin VB.Frame FrameSearchGrid 
      Caption         =   "Search Results:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   0
      TabIndex        =   56
      Top             =   1200
      Width           =   11415
      Begin VB.CommandButton cmdCancelSelect 
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
         Height          =   855
         Left            =   10320
         Picture         =   "frmSession.frx":1D2E
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   5280
         Width           =   975
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   10320
         Picture         =   "frmSession.frx":2170
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   240
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSDGResults 
         Height          =   5895
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   10398
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         TextStyleFixed  =   1
         FocusRect       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   0
      X2              =   752
      Y1              =   235
      Y2              =   235
   End
   Begin VB.Label Label25 
      Caption         =   " :"
      Height          =   255
      Left            =   8520
      TabIndex        =   60
      Top             =   1695
      Width           =   135
   End
End
Attribute VB_Name = "frmEntry"
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
    SSTAB.Visible = True
    UnLockTabs SSTAB
    FrameSearchGrid.Visible = False
    FrameCourseHeader.Visible = True
    FrameCourseHeader.Enabled = True
    FrameSearch.Enabled = True
    FrameButtons.Enabled = True
    FrameGroups.Enabled = True
    FrameChapters.Enabled = True
    FrameOccupations.Enabled = True
    FrameEmailText.Enabled = True
    FrameEmailAddress.Enabled = True
    FrameAddressText.Enabled = True
    FrameAddress.Enabled = True
    cmdAddEmail.Enabled = True
    cmdAddAddress.Enabled = True
    cboGroup.Visible = True
    cmdAddGroup.Visible = True
    cmdDelete.Enabled = True
    cmdDeleteGroup.Visible = True
    txtSearch.Text = ""
    cmdUpdate.Enabled = False
    cmdAdd.Enabled = True
    cmdExit.Enabled = True
    cmdCancel.Enabled = False
    cmdUpdateEmail.Enabled = False
    cmdUpdateAddress.Enabled = False
    cmdSendPEmail.Enabled = True
    cmdAddDues.Enabled = True
    cmdPicture.Enabled = True
    txtAddress1.Locked = True
    txtAddress2.Locked = True
    txtCity.Locked = True
    txtState.Locked = True
    txtZipcode.Locked = True
    txtPhone.Locked = True
    txtCountry.Locked = True
    txtEmail.Locked = True
    chkBus.Enabled = False
    AddNew = False
    EditChange = False
    AddNew2 = False
    EditChange2 = False
    AddNew3 = False
    EditChange3 = False
    AddNew4 = False
    EditChange4 = False
    AddNew5 = False
    EditChange5 = False
    AddNew6 = False
    EditChange6 = False
ElseIf ComVar = 1 Then          ' Add Engaged
    SSTAB.Tab = 0
    LockTabs SSTAB
    EntryImage.Tag = LoadImage(EntryImage, MyPath & "NoImage.Jpg", ShapeBox)
    cmdAdd.Enabled = False
    cmdCancel.Enabled = True
    cmdDelete.Enabled = False
    cmdExit.Enabled = False
    cmdSendPEmail.Enabled = False
    EditCheck cmdUpdate
    AddNew = True
    ClearFrame Me, FrameCourseInformation
    ClearFrame Me, FrameCourseHeader
    ClearFrame Me, FrameSearch
    FrameSearch.Enabled = False
    FrameGroups.Enabled = False
    FrameChapters.Enabled = False
    FrameOccupations.Enabled = False
    MSDGGroups.Visible = False
    chkWidowed.Value = 0
    chkDeceased.Value = 0
    cboGroup.Visible = False
    txtGroupMembershipStatus.Text = "Add Categories After Save"
    txtGroupMembershipStatus.Visible = True
    cmdAddGroup.Visible = False
    cmdDeleteGroup.Visible = False
    DTPLocal.Value = SQLDate
    DTPNational.Value = SQLDate
    DTPLocal.Value = ""
    DTPNational.Value = ""
    txtAddress1.Locked = False
    txtAddress2.Locked = False
    txtCity.Locked = False
    txtState.Locked = False
    txtZipcode.Locked = False
    txtPhone.Locked = False
    txtCountry.Locked = False
    txtEmail.Locked = False
    chkBus.Enabled = True
    txtPrefix.SetFocus
ElseIf ComVar = 2 Then          ' Search Engaged
    SSTAB.Visible = False
    FrameCourseHeader.Visible = False
    FrameSearch.Enabled = False
    FrameSearchGrid.Visible = True
    FrameButtons.Enabled = False
ElseIf ComVar = 3 Then
    LockTabs SSTAB
    cmdAddEmail.Enabled = False
    cmdDelEmail.Enabled = False
    cmdSetPrimaryEmail.Enabled = False
    cmdSendEmail.Enabled = False
    cmdUpdateEmail.Enabled = True
    FrameAddressText.Enabled = False
    FrameAddress.Enabled = False
    FrameCourseHeader.Enabled = False
    FrameSearch.Enabled = False
    FrameButtons.Enabled = False
    FrameEmailText.Visible = True
    txtEEMail.Text = ""
    txtEEMail.SetFocus
    AddNew2 = True
ElseIf ComVar = 4 Then
    LockTabs SSTAB
    cmdAddAddress.Enabled = False
    cmdDelAddress.Enabled = False
    cmdSetPrimaryAddress.Enabled = False
    cmdUpdateAddress.Enabled = True
    FrameEmailText.Enabled = False
    FrameEmailAddress.Enabled = False
    FrameCourseHeader.Enabled = False
    FrameSearch.Enabled = False
    FrameButtons.Enabled = False
    FrameAddressText.Visible = True
    ClearFrame Me, FrameAddressText
    txtSAddress1.SetFocus
    AddNew3 = True
ElseIf ComVar = 5 Then
    LockTabs SSTAB
    cmdAddDues.Enabled = False
    cmdDelDues.Enabled = False
    cmdUpdateDues.Enabled = True
    FrameCourseHeader.Enabled = False
    FrameSearch.Enabled = False
    FrameButtons.Enabled = False
    FrameDuesText.Visible = True
    ClearFrame Me, FrameDuesText
    DTPDues.Value = SQLDate
    txtDues.SetFocus
    AddNew4 = True
End If
End Sub

Private Sub cboEntry_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdRetrieve_Click
End If
End Sub

Private Sub cmdAddAddress_Click()
If UserClick = True Then SetObjects 4
End Sub


Private Sub cmdAddChapter_Click()
If UserClick = True Then
    If Exists("LocalChapter", "EntryID", Trim(Str(FormCurrentValue)), "LocalChapterNumber", Trim(txtLocalChapterNumber.Text)) = True Then
        MsgBox "Entry already exists.", vbCritical
        Exit Sub
    End If
    
    StartRoutine Me
    Dim RS As Recordset
    Set RS = SaveTemplate("LocalChapter", "LocalChapterID", FormCurrentValue)
    
    RS.AddNew
    RS!EntryID = FormCurrentValue
    RS!LocalChapterDate = DTPLocal.Value
    RS!LocalChapterNumber = Trim(txtLocalChapterNumber.Text)
    RS.Update
    
    CloseRS RS
    Set RS = Nothing
    EndRoutine Me
    
    FillChapterGrid FormCurrentValue
End If
End Sub

Private Sub cmdAddDues_Click()
If UserClick = True Then SetObjects 5
End Sub

Private Sub cmdAddEmail_Click()
If UserClick = True Then SetObjects 3
End Sub

Private Sub cmdAddOccupation_Click()
If UserClick = True Then
    If Exists("Occupation", "EntryID", Trim(Str(FormCurrentValue)), "Occupation", Trim(txtOccupationDesc.Text)) = True Then
        MsgBox "Entry already exists.", vbCritical
        Exit Sub
    End If
    
    StartRoutine Me
    Dim RS As Recordset
    Set RS = SaveTemplate("Occupation", "OccupationID", FormCurrentValue)
    
    RS.AddNew
    RS!EntryID = FormCurrentValue
    RS!Occupation = Trim(txtOccupationDesc.Text)
    RS!Classification = Trim(txtOccupationClass.Text)
    RS.Update
    
    CloseRS RS
    Set RS = Nothing
    EndRoutine Me
    
    FillOccupationGrid FormCurrentValue
End If
End Sub

Private Sub cmdDelDues_Click()
If UserClick = True Then
    If MsgBox("Are you sure you want to delete this dues entry?", vbYesNo) = vbYes Then
        DeleteTemplate "Dues", "DuesID", Trim(MSDGDues.TextMatrix(MSDGDues.Row, 0))
        EntryUpdate FormCurrentValue
    End If
End If
End Sub

Private Sub cmdDeleteChapter_Click()
If UserClick = True Then
    If MsgBox("Are you sure you want to delete this chapter?", vbYesNo) = vbYes Then
        DeleteTemplate "LocalChapter", "LocalChapterID", Trim(MSDGChapters.TextMatrix(MSDGChapters.Row, 0))
        FillChapterGrid FormCurrentValue
    End If
End If
End Sub

Private Sub cmdDeleteOccupation_Click()
If UserClick = True Then
    If MsgBox("Are you sure you want to delete this occupation?", vbYesNo) = vbYes Then
        DeleteTemplate "Occupation", "OccupationID", Trim(MSDGOccupations.TextMatrix(MSDGOccupations.Row, 0))
        FillOccupationGrid FormCurrentValue
    End If
End If
End Sub

Private Sub cmdPicture_Click()
Dim PictureFile As String
If UserClick = True Then
    PictureFile = BrowseFile("O", MDIMain.CDBox, "Open Picture File")
    If FileExists(PictureFile) = True Then
        If LoadImage(EntryImage, PictureFile, ShapeBox) = False Then
            ShapeBox.Tag = LoadImage(EntryImage, MyPath & "NoImage.Jpg", ShapeBox)
        Else
            EditCheck cmdUpdate, cmdAdd, cmdCancel
        End If
    End If
End If
End Sub

Private Sub cmdSetPrimaryAddress_Click()
If UserClick = True Then
    SetCommonTableFields "Address", "EntryID", FormCurrentValue, "PreferredAddress", False, True, "AddressID", MSDGAddress.Tag
    EntryUpdate FormCurrentValue
End If
End Sub

Private Sub cmdSetPrimaryEmail_Click()
If UserClick = True Then
    SetCommonTableFields "Email", "EntryID", FormCurrentValue, "PreferredAddress", False, True, "EmailID", MSDGEmailAddress.Tag
    EntryUpdate FormCurrentValue
End If
End Sub

Private Sub cmdAddGroup_Click()
If UserClick = True Then
    If Exists("MembershipGroupAssociation", "EntryID", Trim(Str(FormCurrentValue)), "MembershipGroupID", Trim(cboGroup.ItemData(cboGroup.ListIndex))) = True Then
        MsgBox "Entry is already in this Group.", vbCritical
        Exit Sub
    End If
    
    StartRoutine Me
    Dim RS As Recordset
    Set RS = SaveTemplate("MembershipGroupAssociation", "MembershipGroupAssociationID", FormCurrentValue)
    
    RS.AddNew
    RS!EntryID = FormCurrentValue
    RS!MembershipGroupID = Trim(cboGroup.ItemData(cboGroup.ListIndex))
    RS.Update
    
    CloseRS RS
    Set RS = Nothing
    EndRoutine Me
    
    FillMembershipGrid FormCurrentValue
End If
End Sub

Private Sub cmdDelEmail_Click()
If UserClick = True Then
    If MsgBox("Are you sure you want to delete this email address?", vbYesNo) = vbYes Then
        DeleteTemplate "Email", "EmailID", Trim(MSDGEmailAddress.TextMatrix(MSDGEmailAddress.Row, 0))
        EntryUpdate FormCurrentValue
    End If
End If
End Sub

Private Sub cmdDelAddress_Click()
If UserClick = True Then
 If MsgBox("Are you sure you want to delete this address?", vbYesNo) = vbYes Then
    DeleteTemplate "Address", "AddressID", Trim(MSDGAddress.TextMatrix(MSDGAddress.Row, 0))
    EntryUpdate FormCurrentValue
 End If
End If
End Sub

Private Sub cmdDeleteGroup_Click()
If UserClick = True Then
    If MsgBox("Are you sure you want to delete this group?", vbYesNo) = vbYes Then
        DeleteTemplate "MembershipGroupAssociation", "MembershipGroupAssociationID", Trim(MSDGGroups.TextMatrix(MSDGGroups.Row, 0))
        FillMembershipGrid FormCurrentValue
    End If
End If
End Sub

Private Sub cmdDelete_Click()
If UserClick = True Then
    If MsgBox("Are you sure you want to delete this retiree?", vbYesNo) = vbYes Then
        DeleteTemplate "Entry", "EntryID", FormCurrentValue
        DeleteTemplate "MembershipGroupAssociation", "EntryID", FormCurrentValue
        DeleteTemplate "Address", "EntryID", FormCurrentValue
        DeleteTemplate "Email", "EntryID", FormCurrentValue

        
        DE.rsEntryFullNames.Close
        DE.EntryFullNames
        cboEntry.RowMember = ""
        cboEntry.RowMember = "EntryFullNames"
        cboEntry.Text = ""
        
        EntryUpdate ""
    End If
End If
End Sub

Private Sub cmdSearch_Click()
If UserClick = True And Len(txtSearch.Text) > 2 Then
    StartRoutine Me
        If FillSearchGrid(cboSearch.ItemData(cboSearch.ListIndex)) = True Then
            SetObjects 2
        Else
            MsgBox "Search criteria yielded no results", vbCritical, "Invalid Search"
        End If
    EndRoutine Me
Else
    MsgBox "Search criteria must contain at least 3 characters", vbCritical, "Invalid Search"
End If
End Sub

Private Function FillSearchGrid(OptionFlag As Byte) As Boolean
If OptionFlag = 0 Then
    DE.EntrySearch Trim(txtSearch.Text)
'ElseIf OptionFlag = 1 Then
    'DE.CourseMain 0, "*", Trim(txtSearch.Text) ', "*", "*"
'ElseIf OptionFlag = 2 Then
'    DE.CourseMain 0, "*", "*", Trim(txtSearch.Text), "*"
'ElseIf OptionFlag = 3 Then
'    DE.CourseMain 0, "*", "*", "*", Trim(txtSearch.Text)
Else
    FillSearchGrid = False
    Exit Function
End If

    If Not DE.rsEntrySearch.EOF Then
        Set MSDGResults.DataSource = DE.rsEntrySearch
        MSDGResults.Visible = True
        FillSearchGrid = True
    Else
        FillSearchGrid = False
        MSDGResults.Visible = False
    End If
DE.rsEntrySearch.Close

GridSet MSDGResults, 1

GridFix MSDGResults
End Function

Private Sub cmdCancel_Click()
If UserClick = True Then EntryUpdate FormPassedValue
End Sub

Private Sub cmdCancelSelect_Click()
Call cmdCancel_Click
End Sub

Private Sub cmdSelect_Click()
' Select Attorney From Grid
If UserClick = True Then
    StartRoutine Me
        GridFix MSDGResults
        EntryUpdate Trim(MSDGResults.TextMatrix(MSDGResults.Row, 0))
    EndRoutine Me
End If
End Sub

Private Sub cmdSendEmail_Click()
If UserClick = True And ValidEmail(txtEEMail.Text) = True Then MailTo txtEEMail.Text
End Sub
Private Sub cmdSendPEmail_Click()
If UserClick = True And ValidEmail(txtEmail.Text) = True Then MailTo txtEmail.Text
End Sub

Private Sub cmdUpdate_Click()
If UserClick = True And EditChange = True Then
    If Len(txtFirstName.Text) = 0 Or Len(txtLastName.Text) = 0 Then
        MsgBox "Incomplete Entry, Please complete entry.", vbCritical
        Exit Sub
    End If
    
    StartRoutine Me
        Dim RS As Recordset
        Set RS = SaveTemplate("Entry", "EntryID", FormCurrentValue)
        
        If AddNew = True Then RS.AddNew
                
        RS!PicturePath = EntryImage.Tag
        RS!Active = chkActive.Value
        RS!Deceased = chkDeceased.Value
        RS!Widowed = chkWidowed.Value
        RS!NonMember = chkNonMember.Value
        RS!Prefix = Trim(txtPrefix.Text)
        RS!FirstName = Trim(txtFirstName.Text)
        RS!MiddleName = Trim(txtMiddleName.Text)
        RS!LastName = Trim(txtLastName.Text)
        RS!Suffix = Trim(txtSuffix.Text)
        'RS!LocalChapterDate = Trim(DTPLocal.Value)
        RS!NationalChapterDate = Trim(DTPNational.Value)
        'RS!LocalChapterNumber = Trim(txtLocalChapterNumber.Text)
        RS!NationalChapterNumber = Trim(txtNationalChapterNumber.Text)
        'RS!OccupationDescription = Trim(txtOccupationDesc.Text)
        'RS!OccupationClassification = Trim(txtOccupationClass.Text)
        RS!Comments = Trim(txtComments.Text)
        RS.Update
        
        If AddNew = True Then FormCurrentValue = RS!EntryID
    
        CloseRS RS
        Set RS = Nothing
        
        If AddNew = True Then
            cmdUpdateAddress_Click
            cmdUpdateEmail_Click
        End If
        
        ' RESET Course Combo Box Data to reflect changes made.
            DE.rsEntryFullNames.Close
            DE.EntryFullNames
            cboEntry.RowMember = ""
            cboEntry.RowMember = "EntryFullNames"
            cboEntry.Text = ""
        
    EndRoutine Me
EntryUpdate FormCurrentValue
End If
End Sub



Private Sub cmdUpdateAddress_Click()
Dim Preferred As Boolean
If UserClick = True Or AddNew = True Then
    If Not (Len(txtSAddress1.Text) = 0 Or Len(txtSCity.Text) = 0 Or Len(txtSState.Text) = 0) Then
        If Exists("Address", "EntryID", Trim(Str(FormCurrentValue))) = False Then
            Preferred = True
        Else
            If AddNew3 = True Then
                If MsgBox("Would you like to make this the Preferred Address?", vbYesNo) = vbYes Then
                    Preferred = True
                    SetCommonTableFields "Address", "EntryID", FormCurrentValue, "PreferredAddress", False
                Else
                    Preferred = False
                End If
            End If
        End If
    
        StartRoutine Me
        Dim RS As Recordset
        Set RS = SaveTemplate("Address", "AddressID", MSDGAddress.Tag)
    
        If AddNew3 = True Or AddNew = True Then
            RS.AddNew
            RS!PreferredAddress = Preferred
        End If
        
        RS!EntryID = FormCurrentValue
        RS!Address1 = Trim(txtSAddress1.Text)
        RS!Address2 = Trim(txtSAddress2.Text)
        RS!City = Trim(txtSCity.Text)
        RS!State = Trim(txtSState.Text)
        RS!ZipCode = Trim(txtSZip.Text)
        RS!Country = Trim(txtSCountry.Text)
        RS!Phone = Trim(txtSPhone.Text)
        RS!BusinessAddress = OptBusYes.Value
        RS.Update
    
        CloseRS RS
        Set RS = Nothing
        EndRoutine Me
    
        If AddNew = False Then EntryUpdate FormCurrentValue
    Else
        If AddNew = False Then
            MsgBox "Invalid or Incomplete Address entered, Save Cancelled.", vbCritical
            EntryUpdate FormCurrentValue
        End If
        Exit Sub
    End If
End If
End Sub

Private Sub cmdUpdateDues_Click()
If UserClick = True Then
    If Val(txtDues.Text) > 0 Then
        StartRoutine Me
        Dim RS As Recordset
        Set RS = SaveTemplate("Dues", "DuesID", MSDGDues.Tag)
    
        If AddNew4 = True Then RS.AddNew
        
        RS!EntryID = FormCurrentValue
        RS!Amount = Val(txtDues.Text)
        RS!PaymentDate = DateCheck(DTPDues.Value)
        RS!Comments = Trim(txtDuesComments.Text)
        RS.Update
    
        CloseRS RS
        Set RS = Nothing
        EndRoutine Me
        
        EntryUpdate FormCurrentValue
    Else
        MsgBox "Invalid Dues Amount entered, Save Cancelled.", vbCritical
        EntryUpdate FormCurrentValue
        Exit Sub
    End If
End If
End Sub

Private Sub cmdUpdateEmail_Click()
Dim NewEmailAddress As String
Dim Preferred As Boolean
If UserClick = True Or AddNew = True Then
    NewEmailAddress = Trim(txtEEMail.Text)  'InputBox("Enter Email Address:", "Email Address Add")
    If ValidEmail(NewEmailAddress) = True Then
        If Exists("Email", "EntryID", Trim(Str(FormCurrentValue)), "EmailAddress", "'" & Trim(NewEmailAddress) & "'") = True Then
            MsgBox "Email is already in this Group.", vbCritical
            EntryUpdate FormCurrentValue
            Exit Sub
        End If
        
        If Exists("Email", "EntryID", Trim(Str(FormCurrentValue))) = False Then
            Preferred = True
        Else
            If AddNew2 = True Then
                If MsgBox("Would you like to make this the Preferred Email Address?", vbYesNo) = vbYes Then
                    Preferred = True
                    SetCommonTableFields "Email", "EntryID", FormCurrentValue, "PreferredAddress", False
                Else
                    Preferred = False
                End If
            End If
        End If
    
        StartRoutine Me
        Dim RS As Recordset
        Set RS = SaveTemplate("Email", "EmailID", MSDGEmailAddress.Tag)
    
        If AddNew2 = True Or AddNew = True Then
            RS.AddNew
            RS!PreferredAddress = Preferred
        End If
        
        RS!EntryID = FormCurrentValue
        RS!EmailAddress = LCase(Trim(NewEmailAddress))
        RS.Update
    
        CloseRS RS
        Set RS = Nothing
        EndRoutine Me
    
        If AddNew = False Then EntryUpdate FormCurrentValue
        'FillEmailAddressGrid FormCurrentValue
        If Preferred = True And AddNew = False Then txtEmail.Text = NewEmailAddress
    Else
        If AddNew = False Then
            MsgBox "Invalid or blank Email Address entered, Save Cancelled.", vbCritical
            EntryUpdate FormCurrentValue
        End If
        Exit Sub
    End If
End If
End Sub

Private Sub MSDGResults_DblClick()
Call cmdSelect_Click
End Sub
Private Sub MSDGResults_RowColChange()
GridFix MSDGResults
End Sub

Private Sub Form_Unload(Cancel As Integer)
Xout FormCurrentValue
End Sub

Private Sub cmdExit_Click()
If UserClick = True Then CloseWindow
End Sub

Private Sub Form_Load()
cboGroup.Tag = FillTemplate(cboGroup, "MembershipGroup", "Description", "MembershipGroupID")
cboGroup.ListIndex = 0

cboSearch.ListIndex = 0
SSTAB.Tab = 0

EntryUpdate FormPassedValue
End Sub

Private Sub EntryUpdate(ComVar As String)
StartRoutine Me
    If ComVar = "" Then ComVar = FirstRecord("Entry", "EntryID", "LastName Asc,FirstName Asc")
        
    FormCurrentValue = ComVar
    
    DE.EntryMain ComVar
    If DE.rsEntryMain.EOF = False Then
        txtPrefix.Text = "" & Trim(DE.rsEntryMain!Prefix)
        txtFirstName.Text = "" & Trim(DE.rsEntryMain!FirstName)
        txtMiddleName.Text = "" & Trim(DE.rsEntryMain!MiddleName)
        txtLastName.Text = "" & Trim(DE.rsEntryMain!LastName)
        txtSuffix.Text = "" & Trim(DE.rsEntryMain!Suffix)
        txtComments.Text = "" & Trim(DE.rsEntryMain!Comments)
        txtAddress1.Text = "" & Trim(DE.rsEntryMain!Address1)
        txtAddress2.Text = "" & Trim(DE.rsEntryMain!Address2)
        txtCity.Text = "" & Trim(DE.rsEntryMain!City)
        txtState.Text = "" & Trim(DE.rsEntryMain!State)
        txtZipcode.Text = "" & Trim(DE.rsEntryMain!ZipCode)
        txtCountry.Text = "" & Trim(DE.rsEntryMain!Country)
        txtPhone.Text = PhoneNumberFormat("" & DE.rsEntryMain!Phone, 1)
        'txtLocalChapterNumber.Text = "" & Trim(DE.rsEntryMain!LocalChapterNumber)
        txtNationalChapterNumber.Text = "" & Trim(DE.rsEntryMain!NationalChapterNumber)
        txtEmail.Text = "" & Trim(DE.rsEntryMain!EmailAddress)
        'txtOccupationDesc.Text = "" & Trim(DE.rsEntryMain!OccupationDescription)
        'txtOccupationClass.Text = "" & Trim(DE.rsEntryMain!OccupationClassification)
        'DTPLocal.Value = "" & Trim(DE.rsEntryMain!LocalChapterDate)
        DTPNational.Value = "" & Trim(DE.rsEntryMain!NationalChapterDate)
        chkActive.Value = Abs(CInt(DE.rsEntryMain!Active))
        chkDeceased.Value = Abs(CInt(DE.rsEntryMain!Deceased))
        chkWidowed.Value = Abs(CInt(DE.rsEntryMain!Widowed))
        chkNonMember.Value = Abs(CInt(DE.rsEntryMain!NonMember))
        chkBus.Value = Abs(CInt(DE.rsEntryMain!BusinessAddress))
        Me.Caption = "Retirees   -   Total Entries: " & Trim(Str("" & DE.rsEntryFullNames.RecordCount))
        cboEntry.Text = Trim("" & DE.rsEntryMain!FullName)
        
        If LoadImage(EntryImage, Trim(DE.rsEntryMain!PicturePath), ShapeBox) = False Then
            ShapeBox.Tag = LoadImage(EntryImage, MyPath & "NoImage.Jpg", ShapeBox)
        End If
    End If
    CloseRS DE.rsEntryMain

    FillMembershipGrid FormCurrentValue
    FillAddressGrid FormCurrentValue
    FillEmailAddressGrid FormCurrentValue
    FillDuesGrid FormCurrentValue
    FillChapterGrid FormCurrentValue
    FillOccupationGrid FormCurrentValue
    
    SetObjects 0
EndRoutine Me
End Sub

Private Sub cmdRetrieve_Click()
If UserClick = True And cboEntry.SelectedItem <> cboEntry.Tag Then
    StartRoutine Me
    cboEntry.Tag = cboEntry.SelectedItem
    DE.rsEntryFullNames.Move (Val(cboEntry.Tag) - 1), adBookmarkFirst
    EntryUpdate Trim(DE.rsEntryFullNames!EntryID)
End If
End Sub

Private Sub GridSet(GridObj As MSHFlexGrid, Optional GridNumber As Byte)
With GridObj
  If GridNumber = 0 Then
    .FormatString = "<|<|<|"
    .Cols = 3
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 7500
    .TextMatrix(0, 2) = "Group Membership"
    ElseIf GridNumber = 1 Then
    .FormatString = "<|<|<|<|<|"
    .Cols = 5
    .ColWidth(0) = 0
    .ColWidth(1) = 3000
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .TextMatrix(0, 1) = "Name"
    .TextMatrix(0, 2) = "Active"
    .TextMatrix(0, 3) = "NatNum"
    .TextMatrix(0, 4) = "NatDate"
    ElseIf GridNumber = 2 Then
    .FormatString = "<|<|<|<|"
    .Cols = 4
    .ColWidth(0) = 0
    .ColWidth(1) = 750
    .ColWidth(2) = 4500
    .ColWidth(3) = 0
    .TextMatrix(0, 1) = "Code"
    .TextMatrix(0, 2) = "Category"
    ElseIf GridNumber = 3 Then
    .FormatString = "<|<|<|<|<|<|<|<|<|"
    .Cols = 9
    .ColWidth(0) = 0
    .ColWidth(1) = 1200
    .ColWidth(2) = 800
    .ColWidth(3) = 900
    .ColWidth(4) = 500
    .ColWidth(5) = 725
    .ColWidth(6) = 675
    .ColWidth(7) = 1200
    .ColWidth(8) = 700
    .TextMatrix(0, 1) = "Address1"
    .TextMatrix(0, 2) = "Address2"
    .TextMatrix(0, 3) = "City"
    .TextMatrix(0, 4) = "State"
    .TextMatrix(0, 5) = "ZipCode"
    .TextMatrix(0, 6) = "Country"
    .TextMatrix(0, 7) = "Phone"
    ElseIf GridNumber = 4 Then
    .FormatString = "<|<|<|"
    .Cols = 3
    .ColWidth(0) = 0
    .ColWidth(1) = 3050
    .ColWidth(2) = 0
    .TextMatrix(0, 1) = "EmailAdddress"
    .TextMatrix(0, 2) = "Primary"
    ElseIf GridNumber = 5 Then
    .FormatString = "<|<|<|<|<|"
    .Cols = 5
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 2000
    .ColWidth(3) = 1100
    .ColWidth(4) = 6900
    .TextMatrix(0, 2) = "Dues Paid"
    .TextMatrix(0, 3) = "Date Paid"
    .TextMatrix(0, 4) = "Comments"
    ElseIf GridNumber = 6 Then
    .FormatString = "<|<|<|"
    .Cols = 3
    .ColWidth(0) = 0
    .ColWidth(1) = 1100
    .ColWidth(2) = 2100
    .TextMatrix(0, 1) = "Date"
    .TextMatrix(0, 2) = "Chapter"
    ElseIf GridNumber = 7 Then
    .FormatString = "<|<|<|"
    .Cols = 3
    .ColWidth(0) = 0
    .ColWidth(1) = 3200
    .ColWidth(2) = 3200
    .TextMatrix(0, 1) = "Occupation"
    .TextMatrix(0, 2) = "Classification"
    End If
  .Row = 1
  .Col = 1
End With
End Sub

Private Sub FillEmailAddressGrid(ByVal ComVar As Long)
DE.EmailMain ComVar
    FrameEmailAddress.Caption = "Email Addresses:" & DE.rsEmailMain.RecordCount
    If Not DE.rsEmailMain.EOF Then
        Set MSDGEmailAddress.DataSource = DE.rsEmailMain
        txtEmailStatus.Visible = False
        MSDGEmailAddress.Visible = True
        cmdDelEmail.Enabled = True
        cmdSendEmail.Enabled = True
        If DE.rsEmailMain.RecordCount > 1 Then
            cmdSetPrimaryEmail.Enabled = True
        Else
            cmdSetPrimaryEmail.Enabled = False
        End If
        FrameEmailText.Visible = True
    Else
        FrameEmailText.Visible = False
        cmdSendEmail.Enabled = False
        cmdDelEmail.Enabled = False
        cmdSetPrimaryEmail.Enabled = False
        txtEmailStatus.Visible = True
        MSDGEmailAddress.Visible = False
    End If
DE.rsEmailMain.Close

GridSet MSDGEmailAddress, 4

MSDGEmailAddress_Click
End Sub

Private Sub MSDGEmailAddress_Click()
StartRoutine Me
    GridFix MSDGEmailAddress
    If MSDGEmailAddress.Row > 0 And MSDGEmailAddress.TextMatrix(MSDGEmailAddress.Row, 0) <> "" Then
        MSDGEmailAddress.Tag = Trim(MSDGEmailAddress.TextMatrix(MSDGEmailAddress.Row, 0))
        txtEEMail.Text = Trim(MSDGEmailAddress.TextMatrix(MSDGEmailAddress.Row, 1))
    End If
EndRoutine Me
End Sub
Private Sub MSDGEmailAddress_DblClick()
MSDGEmailAddress_Click
End Sub
Private Sub MSDGEmailAddress_RowColChange()
MSDGEmailAddress_Click
End Sub

Private Sub FillDuesGrid(ByVal ComVar As Long)
DE.DuesMain ComVar
    FrameDues.Caption = "Dues Paid:" & DE.rsDuesMain.RecordCount
    If Not DE.rsDuesMain.EOF Then
        Set MSDGDues.DataSource = DE.rsDuesMain
        txtDuesStatus.Visible = False
        MSDGDues.Visible = True
        cmdDelDues.Enabled = True
        FrameDuesText.Visible = True
    Else
        FrameDuesText.Visible = False
        cmdDelDues.Enabled = False
        txtDuesStatus.Visible = True
        MSDGDues.Visible = False
    End If
DE.rsDuesMain.Close

GridSet MSDGDues, 5

MSDGDues_Click
End Sub

Private Sub MSDGDues_Click()
StartRoutine Me
    GridFix MSDGDues
    If MSDGDues.Row > 0 And MSDGDues.TextMatrix(MSDGDues.Row, 0) <> "" Then
        MSDGDues.Tag = Trim(MSDGDues.TextMatrix(MSDGDues.Row, 0))
        txtDues.Text = Trim(MSDGDues.TextMatrix(MSDGDues.Row, 2))
        DTPDues.Value = DateCheck(MSDGDues.TextMatrix(MSDGDues.Row, 3))
        txtDuesComments.Text = Trim(MSDGDues.TextMatrix(MSDGDues.Row, 4))
    End If
EndRoutine Me
End Sub
Private Sub MSDGDues_DblClick()
MSDGDues_Click
End Sub
Private Sub MSDGDues_RowColChange()
MSDGDues_Click
End Sub

Private Sub FillAddressGrid(ByVal ComVar As Long)
DE.AddressMain ComVar
    FrameAddress.Caption = "Addresses:" & DE.rsAddressMain.RecordCount
    If Not DE.rsAddressMain.EOF Then
        Set MSDGAddress.DataSource = DE.rsAddressMain
        txtAddressStatus.Visible = False
        MSDGAddress.Visible = True
        cmdDelAddress.Enabled = True
        If DE.rsAddressMain.RecordCount > 1 Then
            cmdSetPrimaryAddress.Enabled = True
        Else
            cmdSetPrimaryAddress.Enabled = False
        End If
        FrameAddressText.Visible = True
    Else
        FrameAddressText.Visible = False
        cmdDelAddress.Enabled = False
        cmdSetPrimaryAddress.Enabled = False
        txtAddressStatus.Visible = True
        MSDGAddress.Visible = False
    End If
DE.rsAddressMain.Close

GridSet MSDGAddress, 3

MSDGAddress_Click
End Sub

Private Sub MSDGAddress_Click()
StartRoutine Me
    GridFix MSDGAddress
    If MSDGAddress.Row > 0 And MSDGAddress.TextMatrix(MSDGAddress.Row, 0) <> "" Then
        MSDGAddress.Tag = Trim(MSDGAddress.TextMatrix(MSDGAddress.Row, 0))
        txtSAddress1.Text = Trim(MSDGAddress.TextMatrix(MSDGAddress.Row, 1))
        txtSAddress2.Text = Trim(MSDGAddress.TextMatrix(MSDGAddress.Row, 2))
        txtSCity.Text = Trim(MSDGAddress.TextMatrix(MSDGAddress.Row, 3))
        txtSState.Text = Trim(MSDGAddress.TextMatrix(MSDGAddress.Row, 4))
        txtSZip.Text = Trim(MSDGAddress.TextMatrix(MSDGAddress.Row, 5))
        txtSCountry.Text = Trim(MSDGAddress.TextMatrix(MSDGAddress.Row, 6))
        txtSPhone.Text = PhoneNumberFormat((MSDGAddress.TextMatrix(MSDGAddress.Row, 7)), 1)
        OptBusYes.Value = Trim(MSDGAddress.TextMatrix(MSDGAddress.Row, 8))
        OptBusNo.Value = Not (OptBusYes.Value)
    End If
EndRoutine Me
End Sub
Private Sub MSDGAddress_DblClick()
MSDGAddress_Click
End Sub
Private Sub MSDGAddress_RowColChange()
MSDGAddress_Click
End Sub

Private Sub txtFirstName_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
End Sub

Private Sub txtMiddleName_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
End Sub


Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And UserClick = True Then KeyAscii = 0: cmdSearch_Click
End Sub

Private Sub FillMembershipGrid(ByVal ComVar As Long)
DE.MembershipMain ComVar
    If Not DE.rsMembershipMain.EOF Then
        Set MSDGGroups.DataSource = DE.rsMembershipMain
        txtGroupMembershipStatus.Visible = False
        MSDGGroups.Visible = True
        cmdDeleteGroup.Enabled = True
    Else
        cmdDeleteGroup.Enabled = False
        txtGroupMembershipStatus.Text = "No Categories Found"
        txtGroupMembershipStatus.Visible = True
        MSDGGroups.Visible = False
    End If
DE.rsMembershipMain.Close

GridSet MSDGGroups, 0

MSDGGroups_Click
End Sub

Private Sub MSDGGroups_Click()
StartRoutine Me
    GridFix MSDGGroups
If MSDGGroups.Row > 0 And MSDGGroups.TextMatrix(MSDGGroups.Row, 0) <> "" Then
    MSDGGroups.Tag = Trim(MSDGGroups.TextMatrix(MSDGGroups.Row, 0))
    cboGroup.Text = Trim(MSDGGroups.TextMatrix(MSDGGroups.Row, 2))
End If
EndRoutine Me
End Sub
Private Sub MSDGGroups_DblClick()
MSDGGroups_Click
End Sub
Private Sub MSDGGroups_RowColChange()
MSDGGroups_Click
End Sub

Private Sub FillChapterGrid(ByVal ComVar As Long)
DE.ChapterMain ComVar
    If Not DE.rsChapterMain.EOF Then
        Set MSDGChapters.DataSource = DE.rsChapterMain
        txtChapterStatus.Visible = False
        MSDGChapters.Visible = True
        cmdDeleteChapter.Enabled = True
    Else
        cmdDeleteGroup.Enabled = False
        txtChapterStatus.Text = "No Chapters Found"
        txtChapterStatus.Visible = True
        MSDGChapters.Visible = False
    End If
DE.rsChapterMain.Close

GridSet MSDGChapters, 6

MSDGChapters_Click
End Sub

Private Sub MSDGChapters_Click()
StartRoutine Me
    GridFix MSDGChapters
If MSDGChapters.Row > 0 And MSDGChapters.TextMatrix(MSDGChapters.Row, 0) <> "" Then
    MSDGChapters.Tag = Trim(MSDGChapters.TextMatrix(MSDGChapters.Row, 0))
    DTPLocal.Value = Trim(MSDGChapters.TextMatrix(MSDGChapters.Row, 1))
    txtLocalChapterNumber.Text = Trim(MSDGChapters.TextMatrix(MSDGChapters.Row, 2))
End If
EndRoutine Me
End Sub
Private Sub MSDGChapters_DblClick()
MSDGChapters_Click
End Sub
Private Sub MSDGChapters_RowColChange()
MSDGChapters_Click
End Sub

Private Sub FillOccupationGrid(ByVal ComVar As Long)
DE.OccupationMain ComVar
    If Not DE.rsOccupationMain.EOF Then
        Set MSDGOccupations.DataSource = DE.rsOccupationMain
        txtOccupationStatus.Visible = False
        MSDGOccupations.Visible = True
        cmdDeleteOccupation.Enabled = True
    Else
        cmdDeleteOccupation.Enabled = False
        txtOccupationStatus.Text = "No Occupations Found"
        txtOccupationStatus.Visible = True
        MSDGOccupations.Visible = False
    End If
DE.rsOccupationMain.Close

GridSet MSDGOccupations, 7

MSDGOccupations_Click
End Sub

Private Sub MSDGOccupations_Click()
StartRoutine Me
    GridFix MSDGOccupations
If MSDGOccupations.Row > 0 And MSDGOccupations.TextMatrix(MSDGOccupations.Row, 0) <> "" Then
    MSDGOccupations.Tag = Trim(MSDGOccupations.TextMatrix(MSDGOccupations.Row, 0))
    txtOccupationDesc.Text = Trim(MSDGOccupations.TextMatrix(MSDGOccupations.Row, 1))
    txtOccupationClass.Text = Trim(MSDGOccupations.TextMatrix(MSDGOccupations.Row, 2))
End If
EndRoutine Me
End Sub
Private Sub MSDGOccupations_DblClick()
MSDGOccupations_Click
End Sub
Private Sub MSDGOccupations_RowColChange()
MSDGOccupations_Click
End Sub

'------------------------------------EDITCHECKS----------------------------------------
Private Sub txtAddress1_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
txtSAddress1.Text = txtAddress1.Text
End Sub
Private Sub txtAddress2_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
txtSAddress2.Text = txtAddress2.Text
End Sub
Private Sub txtCity_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
txtSCity.Text = txtCity.Text
End Sub
Private Sub txtComments_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
End Sub
Private Sub txtCountry_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
txtSCountry.Text = txtCountry.Text
End Sub
Private Sub txtCourseDescription_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
End Sub
Private Sub txtSPhone_Change()
EditCheck2 cmdUpdateAddress, cmdAddAddress
End Sub
Private Sub txtSState_Change()
EditCheck2 cmdUpdateAddress, cmdAddAddress
End Sub
Private Sub txtEmail_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
txtEEMail.Text = txtEmail.Text
End Sub
Private Sub txtSuffix_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
End Sub
Private Sub txtZipcode_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
txtSZip.Text = txtZipcode.Text
End Sub
Private Sub txtState_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
txtSState.Text = txtState.Text
End Sub
Private Sub txtPhone_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
txtSPhone.Text = txtPhone.Text
End Sub
Private Sub txtSZip_Change()
EditCheck2 cmdUpdateAddress, cmdAddAddress
End Sub
Private Sub txtEEMail_Change()
EditCheck2 cmdUpdateEmail, cmdAddEmail
End Sub
Private Sub txtSAddress_Change()
EditCheck3 cmdUpdateAddress, cmdAddAddress
End Sub
Private Sub txtSAddress2_Change()
EditCheck2 cmdUpdateAddress, cmdAddAddress
End Sub
Private Sub txtSCity_Change()
EditCheck2 cmdUpdateAddress, cmdAddAddress
End Sub
Private Sub txtSCountry_Change()
EditCheck2 cmdUpdateAddress, cmdAddAddress
End Sub
Private Sub OptBusNo_Click()
EditCheck3 cmdUpdateAddress, cmdAddAddress
End Sub
Private Sub OptBusYes_Click()
EditCheck3 cmdUpdateAddress, cmdAddAddress
End Sub
Private Sub txtNationalChapterNumber_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
End Sub
Private Sub txtLocalChapterNumber_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
End Sub
Private Sub DTPLocal_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
End Sub
Private Sub DTPNational_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
End Sub
Private Sub txtLastName_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
End Sub
Private Sub chkActive_Click()
EditCheck cmdUpdate, cmdAdd, cmdCancel
If UserClick = True Then
    StartRoutine Me
    chkDeceased.Value = 0
    EndRoutine Me
End If
End Sub
Private Sub chkNonMember_Click()
EditCheck cmdUpdate, cmdAdd, cmdCancel
If UserClick = True Then
    StartRoutine Me
    chkDeceased.Value = 0
    EndRoutine Me
End If
End Sub
Private Sub chkWidowed_Click()
EditCheck cmdUpdate, cmdAdd, cmdCancel
If UserClick = True Then
    StartRoutine Me
    chkDeceased.Value = 0
    EndRoutine Me
End If
End Sub
Private Sub chkDeceased_Click()
EditCheck cmdUpdate, cmdAdd, cmdCancel
If UserClick = True Then
    StartRoutine Me
        chkWidowed.Value = 0
        chkNonMember.Value = 0
        chkActive.Value = 0
    EndRoutine Me
End If
End Sub
Private Sub txtOccupationClass_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
End Sub
Private Sub txtOccupationDesc_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
End Sub
Private Sub chkBus_Click()
EditCheck cmdUpdate, cmdAdd, cmdCancel
OptBusYes.Value = CBool(chkBus.Value)
'OptBusNo.Value = Not CBool(chkBus.Value)
End Sub
Private Sub txtDuesComments_Change()
EditCheck4 cmdUpdateDues, cmdAddDues
End Sub
Private Sub txtDues_Change()
EditCheck4 cmdUpdateDues, cmdAddDues
End Sub
Private Sub DTPDues_Change()
EditCheck4 cmdUpdateDues, cmdAddDues
End Sub
Private Sub txtPrefix_Change()
EditCheck cmdUpdate, cmdAdd, cmdCancel
End Sub
'---------------------------------------------------------------------------------
