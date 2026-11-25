VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWorkOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Work Order"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   581
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   667
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
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   56
      Text            =   "No Work Orders Found"
      Top             =   6720
      Width           =   3255
   End
   Begin VB.CommandButton cmdAddItem 
      Caption         =   "Add"
      Height          =   435
      Left            =   9240
      TabIndex        =   26
      Top             =   3840
      Width           =   615
   End
   Begin VB.Frame FrameMain 
      Height          =   2175
      Left            =   0
      TabIndex        =   28
      Top             =   1200
      Width           =   9975
      Begin VB.CommandButton cmdPickCus 
         Height          =   285
         Left            =   9480
         Picture         =   "frmWorkOrder.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdPickEmp 
         Height          =   285
         Left            =   3480
         Picture         =   "frmWorkOrder.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   960
         Width           =   375
      End
      Begin VB.ComboBox cboCustomer 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   960
         Width           =   4335
      End
      Begin VB.ComboBox cboEmployee 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtDescription 
         Height          =   285
         Left            =   960
         MaxLength       =   255
         TabIndex        =   11
         Top             =   600
         Width           =   8895
      End
      Begin VB.TextBox txtComments 
         Height          =   735
         Left            =   960
         MaxLength       =   1024
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   1320
         Width           =   8895
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
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPCloseDate 
         Height          =   315
         Left            =   8400
         TabIndex        =   10
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "MM/dd/yyy"
         Format          =   119209987
         CurrentDate     =   36706
      End
      Begin MSComCtl2.DTPicker DTPOpenDate 
         Height          =   315
         Left            =   3240
         TabIndex        =   8
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MM/dd/yyy"
         Format          =   119209987
         CurrentDate     =   27908
      End
      Begin MSComCtl2.DTPicker DTPInvoiceDate 
         Height          =   315
         Left            =   5760
         TabIndex        =   9
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "MM/dd/yyy"
         Format          =   119209987
         CurrentDate     =   36706
      End
      Begin VB.Label Label8 
         Caption         =   "Customer:"
         Height          =   255
         Left            =   4320
         TabIndex        =   37
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Employee:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Invoiced:"
         Height          =   255
         Left            =   5040
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Closed:"
         Height          =   255
         Left            =   7800
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Opened:"
         Height          =   255
         Left            =   2520
         TabIndex        =   33
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Title:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Comments:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "ID Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   855
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
      Width           =   9975
      Begin VB.CommandButton cmdReport 
         Caption         =   "Invoice"
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
         Left            =   7560
         Picture         =   "frmWorkOrder.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Pending"
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
         Left            =   5040
         Picture         =   "frmWorkOrder.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   59
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
         Left            =   3720
         Picture         =   "frmWorkOrder.frx":0A88
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
         Left            =   9000
         Picture         =   "frmWorkOrder.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   2520
         Picture         =   "frmWorkOrder.frx":130C
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
         Left            =   1320
         Picture         =   "frmWorkOrder.frx":174E
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
         Picture         =   "frmWorkOrder.frx":1B90
         Style           =   1  'Graphical
         TabIndex        =   1
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
         Left            =   6240
         Picture         =   "frmWorkOrder.frx":22D2
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Line Line10 
         X1              =   8760
         X2              =   8760
         Y1              =   960
         Y2              =   240
      End
      Begin VB.Line Line9 
         X1              =   7320
         X2              =   7320
         Y1              =   960
         Y2              =   240
      End
      Begin VB.Line Line8 
         X1              =   4800
         X2              =   4800
         Y1              =   960
         Y2              =   240
      End
   End
   Begin VB.Frame FrameDetail 
      Height          =   5175
      Left            =   0
      TabIndex        =   32
      Top             =   3525
      Width           =   9975
      Begin VB.CommandButton cmdDelItem 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   435
         Left            =   9240
         TabIndex        =   62
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdateItem 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   435
         Left            =   9240
         TabIndex        =   61
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtSubTotal 
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
         Left            =   600
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   57
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txtNote 
         Height          =   765
         Left            =   960
         MaxLength       =   2048
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   1320
         Width           =   8055
      End
      Begin VB.ComboBox cboItemCategory 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmWorkOrder.frx":2714
         Left            =   960
         List            =   "frmWorkOrder.frx":2716
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   600
         Width           =   6255
      End
      Begin VB.ComboBox cboItem 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmWorkOrder.frx":2718
         Left            =   960
         List            =   "frmWorkOrder.frx":271A
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   960
         Width           =   6255
      End
      Begin VB.TextBox txtTotal 
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
         Left            =   8160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   24
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtAmount 
         Height          =   285
         Left            =   8160
         TabIndex        =   23
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtQuantity 
         Height          =   285
         Left            =   8160
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cboEntity 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Width           =   3855
      End
      Begin VB.CommandButton cmdPickEnt 
         Height          =   285
         Left            =   6840
         Picture         =   "frmWorkOrder.frx":271C
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtPTax 
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
         Left            =   3720
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   46
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox txtSTax 
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
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   44
         Top             =   4800
         Width           =   735
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
         Left            =   6960
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   42
         Top             =   4800
         Width           =   975
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
         Left            =   8880
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   40
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txtOwed 
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
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   38
         Top             =   4800
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSDGWork 
         Height          =   2490
         Left            =   120
         TabIndex        =   27
         Top             =   2160
         Visible         =   0   'False
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4392
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
      Begin MSComCtl2.DTPicker DTPWorkDate 
         Height          =   315
         Left            =   960
         TabIndex        =   17
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MM/dd/yyy"
         Format          =   119209987
         CurrentDate     =   27908
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Sub:"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   4800
         Width           =   375
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000003&
         X1              =   9120
         X2              =   9120
         Y1              =   2160
         Y2              =   120
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000003&
         X1              =   7320
         X2              =   9120
         Y1              =   1290
         Y2              =   1290
      End
      Begin VB.Label Label17 
         Caption         =   "Notes:"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   1320
         Width           =   855
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000003&
         X1              =   7320
         X2              =   7320
         Y1              =   120
         Y2              =   1290
      End
      Begin VB.Label Label16 
         Caption         =   "Category:"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "Item:"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "Total:"
         Height          =   255
         Left            =   7560
         TabIndex        =   52
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "Quantity:"
         Height          =   255
         Left            =   7440
         TabIndex        =   51
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Amount:"
         Height          =   255
         Left            =   7440
         TabIndex        =   50
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Item Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Entity:"
         Height          =   255
         Left            =   2520
         TabIndex        =   48
         Top             =   240
         Width           =   615
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         X1              =   8040
         X2              =   8040
         Y1              =   4680
         Y2              =   5160
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Parish:"
         Height          =   255
         Left            =   3120
         TabIndex        =   47
         Top             =   4800
         Width           =   495
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "State:"
         Height          =   255
         Left            =   1680
         TabIndex        =   45
         Top             =   4800
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Balance:"
         Height          =   255
         Left            =   8040
         TabIndex        =   43
         Top             =   4800
         Width           =   735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         X1              =   4560
         X2              =   4560
         Y1              =   4680
         Y2              =   5160
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Paid:"
         Height          =   255
         Left            =   6360
         TabIndex        =   41
         Top             =   4800
         Width           =   495
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         Caption         =   "Owed:"
         Height          =   255
         Left            =   4680
         TabIndex        =   39
         Top             =   4800
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   9840
         Y1              =   4680
         Y2              =   4680
      End
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   664
      Y1              =   232
      Y2              =   232
   End
End
Attribute VB_Name = "frmWorkOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FormCurrentValue As Long, AddNew As Boolean, EditChange As Boolean, AddNew2 As Boolean, EditChange2 As Boolean


Private Sub cmdAdd_Click()
If UserClick = True Then SetObjects 1
End Sub
Private Sub cmdAddItem_Click()
If UserClick = True Then SetObjects 2
End Sub
Private Sub cmdCancel_Click()
If UserClick = True Then WOLoad
End Sub

Private Sub cmdExit_Click()
If UserClick = True Then Unload Me
End Sub

Private Sub cmdReport_Click()
StartRoutine Me
    Dim RS As ADODB.Recordset
    Set RS = RSTemplate("", "", 0, "Exec InvoiceReport " & FormCurrentValue & ", 0")
    
    frmReport.ViewReport "Invoice", RS, , False
    LoadForm frmReport, 1
    
    CloseRS RS
    Set RS = Nothing
EndRoutine Me
End Sub


Private Sub Form_Activate()
DTPOpenDate.SetFocus
End Sub

Private Sub Form_Load()
cboEmployee.Tag = FillBox(cboEmployee, "", "ListItem", "ListID", "Exec EntityList 1")
cboEmployee.ListIndex = 0

cboCustomer.Tag = FillBox(cboCustomer, "", "ListItem", "ListID", "Exec EntityList 2")
cboCustomer.ListIndex = 0

cboEntity.Tag = FillBox(cboEntity, "", "ListItem", "ListID", "Exec EntityList 3")
cboEntity.ListIndex = 0

cboItemCategory.Tag = FillBox(cboItemCategory, "", "ListItem", "ListID", "Exec ItemList 1")
cboItemCategory.ListIndex = 0

'cboItem.Tag = FillBox(cboItem, "", "ListItem", "ListID", "Exec ItemList 2")
'cboItem.ListIndex = 0

FormCurrentValue = FormPassedValue
FormPassedValue = 0

WOLoad
End Sub

Private Sub WOLoad(Optional PrevPos As Long = 0)
If FormCurrentValue = 0 Then FormCurrentValue = Val(TransTemplate("Select Top 1 WorkOrderID From WorkOrder Order By CloseDate Asc, WorkOrderID Desc", "WorkOrderID", "0"))

Dim RS As ADODB.Recordset
Set RS = RSTemplate("", "", 0, "Exec WorkOrderMain " & Str(FormCurrentValue))

If Not RS.EOF Then
    ' Pre-populate date fields
    DTPInvoiceDate.Value = Date$
    DTPCloseDate.Value = Date$
    ' ************************

    txtID.Text = RS!WorkOrderNumber
    DTPOpenDate.Value = Trim("" & RS!OpenDate)
    DTPInvoiceDate.Value = Trim("" & RS!InvoiceDate)
    DTPCloseDate.Value = Trim("" & RS!CloseDate)
    txtDescription.Text = Trim("" & RS!WorkOrderDesc)
    cboEmployee.Tag = FindID(cboEmployee, Val("" & RS!EmpID))
    cboCustomer.Tag = FindID(cboCustomer, Val("" & RS!CusID))
    
    txtComments.Text = Trim("" & RS!Comments)
    
    txtOwed.Text = Trim("" & RS!Owed)
    txtPaid.Text = Trim("" & RS!Paid)
    txtSTax.Text = Trim("" & RS!STax)
    txtPTax.Text = Trim("" & RS!PTax)
    'txtBTax.Text = Trim("" & RS!BTax)
    txtSubTotal.Text = Trim("" & RS!SubTotal)
    txtBalance.Text = Trim("" & RS!Balance)
    txtBalance.Tag = txtBalance.Text
    
    Select Case Val(txtBalance.Text)
    Case Is > 0
        txtBalance.BackColor = RGB(255, 175, 175)
    Case Is = 0
        txtBalance.BackColor = RGB(175, 255, 175)
    Case Else
        txtBalance.BackColor = &H80000018
    End Select
End If

FillWOGrid FormCurrentValue, PrevPos

CloseRS RS
Set RS = Nothing

SetObjects 0
End Sub

Private Sub cmdUpdate_Click()
If UserClick = True And EditChange = True Then

StartRoutine Me
        Dim RS As ADODB.Recordset
        Set RS = RSTemplate("WorkOrder", "WorkOrderID", FormCurrentValue)
    
        If AddNew = True Then RS.AddNew
                
        RS!EntityID = Trim(cboCustomer.ItemData(cboCustomer.ListIndex))
        RS!ToEntityID = Trim(cboEmployee.ItemData(cboEmployee.ListIndex))
        RS!OpenDate = DTPOpenDate.Value
        RS!InvoiceDate = DTPInvoiceDate.Value
        RS!CloseDate = DTPCloseDate.Value
        RS!Description = Trim(txtDescription.Text)
        RS!Comments = Trim(txtComments.Text)
        
        RS.Update
        
        If AddNew = True Then FormCurrentValue = RS!WorkOrderID
    
        CloseRS RS
        
        Set RS = RSTemplate("", "", 0, "Exec EastTreeUpdater " & FormCurrentValue)
        Set RS = Nothing
    EndRoutine Me
    
    WOLoad
End If
End Sub


Private Sub SetObjects(ComVar As Byte)
Select Case ComVar
Case 0  ' Normal
    'UnLockTabs SSTab, SSTab.Tab
    FrameMain.Enabled = True
    'FrameDetail.Visible = True
    cmdSearch.Enabled = True
    cmdOpen.Enabled = True
    cmdAdd.Enabled = True
    cmdCancel.Enabled = False
    cmdUpdate.Enabled = False
    cmdExit.Enabled = True
    cmdAddItem.Visible = True
    'cmdDelete.Enabled = True
    
    AddEditFalse
Case 1  ' Add Work Order
    'LockTabs SSTab, 0
    ECheck
    AddNew = True
    ClearFrame Me, FrameMain
    FrameDetail.Visible = False
    DTPOpenDate.Value = Date$
    'DTPInvoiceDate.Value = Date$
    DTPCloseDate.Value = ""
    cmdAddItem.Visible = False
    
    DTPOpenDate.SetFocus
Case 2  ' Add Item
    ECheck2
    AddNew2 = True
    ClearFrame Me, FrameDetail
    FrameDetail.Visible = True
    FrameDetail.Enabled = True
    MSDGWork.Visible = False
    txtStatusWork.Visible = False
    FrameMain.Enabled = False
        
    cboEntity.Tag = FindID(cboEntity, Val(cboEmployee.ItemData(cboEmployee.ListIndex)))
    
    DTPWorkDate.Value = Date$
    DTPWorkDate.SetFocus
End Select
End Sub

Private Sub cmdSearch_Click()
If UserClick = True Then
    cmdSearch.Tag = PickBox(7)
    
    If Val(cmdSearch.Tag) > 0 Then FormCurrentValue = Val(cmdSearch.Tag)
    WOLoad
End If
End Sub

Private Sub cmdOpen_Click()
If UserClick = True Then
    cmdSearch.Tag = PickBox(7, "[OpenOnly]")
    
    If Val(cmdSearch.Tag) > 0 Then FormCurrentValue = Val(cmdSearch.Tag)
    WOLoad
End If
End Sub

Private Sub cmdUpdateItem_Click()
If UserClick = True And EditChange2 = True Then

StartRoutine Me
        Dim RS As ADODB.Recordset
        Set RS = RSTemplate("WorkDetail", "WorkDetailID", Val(MSDGWork.Tag))
    
        If AddNew2 = True Then RS.AddNew
                
        RS!WorkOrderID = FormCurrentValue
        RS!WorkDate = DTPWorkDate.Value
        RS!WorkItemID = Val(cboItem.ItemData(cboItem.ListIndex))
        RS!EntityID = Val(cboEntity.ItemData(cboEntity.ListIndex))
        RS!FromEntityID = Val(cboCustomer.ItemData(cboCustomer.ListIndex))
        RS!Description = Trim(txtNote.Text)
        RS!Quantity = Val(txtQuantity.Text)
        RS!Amount = Val(txtAmount.Text)
        RS!OAmount = Val(txtAmount.Text)
        
        RS.Update
        CloseRS RS
        
        Set RS = RSTemplate("", "", 0, "Exec EastTreeUpdater " & FormCurrentValue)
        Set RS = Nothing
    EndRoutine Me
    
    If AddNew2 = True Then
        WOLoad
    Else
        WOLoad MSDGWork.Row
    End If
    
    
End If
End Sub

Private Sub cmdDelItem_Click()
If UserClick = True Then
 If MSDGWork.Visible = True Then
    If MsgBox("Are you sure you want to delete this work order item?", vbYesNo) = vbYes Then
        DeleteTemplate "WorkDetail", "WorkDetailID", Val(MSDGWork.TextMatrix(MSDGWork.Row, 0))
        
        Dim RS As ADODB.Recordset
        Set RS = RSTemplate("", "", 0, "Exec EastTreeUpdater " & FormCurrentValue)
        Set RS = Nothing
        
        WOLoad
    End If
 End If
End If
End Sub

' ************************************** Grids ************************************
Private Sub FillWOGrid(ComID As Long, Optional PrevPos As Long = 0)
StartRoutine Me
    If FillGrid("Exec WorkDetailMain " & Str(ComID), MSDGWork, "", txtStatusWork) = True Then
        With MSDGWork
            .ColWidth(2) = 1000      ' Date
            .ColWidth(3) = 1300      ' EntityTo
            .ColWidth(4) = 2200      ' Item
            .ColWidth(5) = 2700      ' Notes
            .ColWidth(6) = 675       ' Quantity
            .ColWidth(7) = 750       ' Amount
            .ColWidth(8) = 750       ' Total
            .Row = 1
            .Col = 1
        End With
        
        If PrevPos > 0 And PrevPos < MSDGWork.Rows Then MSDGWork.Row = PrevPos

        cmdAddItem.Enabled = True
        cmdDelItem.Enabled = True
        cmdUpdateItem.Enabled = False
        FrameDetail.Visible = True
        FrameDetail.Enabled = True
    Else
        txtStatusWork.Text = "No Item Details Exist"
        cmdAddItem.Enabled = True
        cmdDelItem.Enabled = False
        cmdUpdateItem.Enabled = False
        FrameDetail.Visible = False
        FrameDetail.Enabled = False
    End If
    
    MSDGWork_Click
EndRoutine Me
End Sub


Private Sub MSDGWork_Click()
StartRoutine Me
GridFix MSDGWork: DoEvents

If MSDGWork.Row > 0 And MSDGWork.TextMatrix(MSDGWork.Row, 0) <> "" Then
    MSDGWork.Tag = Trim(MSDGWork.TextMatrix(MSDGWork.Row, 0))
    cboEntity.Tag = FindID(cboEntity, Val(MSDGWork.TextMatrix(MSDGWork.Row, 1)))
    DTPWorkDate.Value = Trim(MSDGWork.TextMatrix(MSDGWork.Row, 2))
    txtQuantity.Text = Val(MSDGWork.TextMatrix(MSDGWork.Row, 6))
    txtAmount.Text = Val(MSDGWork.TextMatrix(MSDGWork.Row, 7))
    txtTotal.Text = Val(MSDGWork.TextMatrix(MSDGWork.Row, 8))
    txtNote.Text = Trim(MSDGWork.TextMatrix(MSDGWork.Row, 5))
    cboItemCategory.Tag = FindItem(cboItemCategory, Trim(MSDGWork.TextMatrix(MSDGWork.Row, 9)))
    cboItem.Tag = FindItem(cboItem, Trim(MSDGWork.TextMatrix(MSDGWork.Row, 4)))
Else
    ClearFrame Me, FrameDetail
End If

EditChange2 = False
EndRoutine Me
End Sub
Private Sub MSDGWork_DblClick()
MSDGWork_Click
End Sub
Private Sub MSDGWork_RowColChange()
MSDGWork_Click
End Sub


' ************************************** Edit Checks ************************************
Private Sub ECheck()
EditChange = EditCheck(cmdUpdate, cmdAdd, cmdCancel, cmdDelete, cmdExit, cmdSearch, cmdOpen)
End Sub
Private Sub ECheck2()
EditChange2 = EditCheck(cmdUpdateItem, cmdAddItem, cmdCancel, cmdDelItem, cmdExit, cmdSearch, cmdUpdate, cmdAdd, cmdDelete, cmdOpen)
End Sub
Public Sub AddEditFalse()
    AddNew = False
    EditChange = False
    AddNew2 = False
    EditChange2 = False
End Sub

Private Sub cmdPickEmp_Click()
If UserClick = True Then
    cmdPickEmp.Tag = PickBox(1)
    
    If Val(cmdPickEmp.Tag) > 0 Then
        cboEmployee.Tag = FindID(cboEmployee, Val(cmdPickEmp.Tag))
        ECheck
    End If
End If
End Sub
Private Sub cmdPickCus_Click()
If UserClick = True Then
    cmdPickCus.Tag = PickBox(2)
    
    If Val(cmdPickEmp.Tag) > 0 Then
        cmdPickCus.Tag = FindID(cboCustomer, Val(cmdPickCus.Tag))
        ECheck
    End If
End If
End Sub
Private Sub cboCustomer_Click()
ECheck
End Sub
Private Sub cboEmployee_Click()
ECheck
End Sub
Private Sub txtComments_Change()
ECheck
End Sub
Private Sub txtDescription_Change()
ECheck
End Sub
Private Sub DTPOpenDate_Change()
ECheck
End Sub
Private Sub DTPInvoiceDate_Change()
ECheck
End Sub
Private Sub DTPCloseDate_Change()
ECheck
End Sub
'********************************************************************

Private Sub cmdPickEnt_Click()
If UserClick = True Then
    cmdPickEnt.Tag = PickBox(8)
    
    If Val(cmdPickEnt.Tag) > 0 Then
        cboEntity.Tag = FindID(cboEntity, Val(cmdPickEnt.Tag))
        ECheck2
    End If
End If
End Sub

Private Sub cboItem_Click()
ECheck2
If UserClick = True Then ItemAmount Val(cboItemCategory.ItemData(cboItemCategory.ListIndex))
End Sub
Private Sub cboItemCategory_Click()
ECheck2

Dim ItemCategoryID As Long
ItemCategoryID = Val(cboItemCategory.ItemData(cboItemCategory.ListIndex))

txtQuantity.Enabled = (ItemCategoryID <> 2)

cboItem.Tag = FillBox(cboItem, "", "ListItem", "ListID", "Exec ItemList 2," & Str(ItemCategoryID))
If UserClick = True Then ItemAmount ItemCategoryID
End Sub
Private Sub ItemAmount(ItemCategoryID As Long)
txtAmount.Tag = TransTemplate("Select Amount From WorkItem Where WorkItemID = " & Val(cboItem.ItemData(cboItem.ListIndex)), "Amount", "0")
If ItemCategoryID = 2 Then
    txtAmount.Tag = txtBalance.Tag
    cboEntity.ListIndex = 0
End If
If Val(txtAmount.Tag) > 0 Then
    txtQuantity.Text = "1"
    txtAmount.Text = txtAmount.Tag
End If
End Sub

Private Sub cboEntity_Click()
ECheck2
End Sub
Private Sub DTPWorkDate_Change()
ECheck2
End Sub
Private Sub txtNote_Change()
ECheck2
End Sub
Private Sub txtAmount_Change()
ECheck2
If UserClick = True Then txtTotal.Text = Val(txtQuantity.Text) * Val(txtAmount.Text)
End Sub
Private Sub txtQuantity_Change()
ECheck2
If UserClick = True Then txtTotal.Text = Val(txtQuantity.Text) * Val(txtAmount.Text)
End Sub
Private Sub txtAmount_KeyPress(KeyAscii As Integer)
KeyAscii = NumbersOnly(KeyAscii)
End Sub
Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
KeyAscii = NumbersOnly(KeyAscii)
End Sub
