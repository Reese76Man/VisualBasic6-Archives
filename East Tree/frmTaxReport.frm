VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTaxReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax & Earnings Reports"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   Icon            =   "frmTaxReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   6060
   Begin VB.ComboBox cboMonth 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmTaxReport.frx":0442
      Left            =   1200
      List            =   "frmTaxReport.frx":0470
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   960
      Width           =   3735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Preview / Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      Picture         =   "frmTaxReport.frx":04DB
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cboYear 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmTaxReport.frx":091D
      Left            =   1200
      List            =   "frmTaxReport.frx":091F
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   600
      Width           =   3735
   End
   Begin VB.ComboBox cboReport 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmTaxReport.frx":0921
      Left            =   1200
      List            =   "frmTaxReport.frx":092B
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   3735
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
      Left            =   5160
      Picture         =   "frmTaxReport.frx":0954
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DTPFrom 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
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
      Format          =   114556931
      CurrentDate     =   27908
   End
   Begin MSComCtl2.DTPicker DTPTo 
      Height          =   315
      Left            =   3600
      TabIndex        =   6
      Top             =   1320
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
      Format          =   114556931
      CurrentDate     =   27908
   End
   Begin VB.Label Label5 
      Caption         =   "Month:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.Line Line2 
      X1              =   5040
      X2              =   5040
      Y1              =   120
      Y2              =   2040
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5040
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label4 
      Caption         =   "Year:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "To:"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "From:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Report Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmTaxReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrint_Click()
StartRoutine Me
    Dim RS As ADODB.Recordset
    Set RS = RSTemplate("", "", 0, "Exec TaxReport " & cboReport.ItemData(cboReport.ListIndex) & ", '" & DTPFrom.Value & "','" & DTPTo.Value & "'")
    
    Select Case cboReport.ItemData(cboReport.ListIndex)
    Case 1
        cboReport.Tag = "TaxReport"
    Case Else
        cboReport.Tag = "IncTaxReport"
    End Select
    
    frmReport.ViewReport cboReport.Tag, RS, , False
    LoadForm frmReport, 1
    
    CloseRS RS
    Set RS = Nothing
EndRoutine Me
End Sub

Private Sub Form_Activate()
cboReport.SetFocus
End Sub

Private Sub Form_Load()
cboReport.ListIndex = 0
cboMonth.ListIndex = 0

cboYear.Tag = FillBox(cboYear, "", "ListItem", "ListID", "Select Distinct YEAR(WD.WorkDate) as ListID, (Convert(VarChar,YEAR(WD.WorkDate)) + ' Calendar Year') As ListItem From WorkDetail WD Order By YEAR(WD.WorkDate) Desc")
cboYear.AddItem "Custom Dates", 0
cboYear.ListIndex = 1: cboYear_Click
cboMonth.ListIndex = Month(Date$): cboMonth_Click
End Sub

Private Sub cboYear_Click()
If UserClick = True Then
    cboYear.Tag = cboYear.ItemData(cboYear.ListIndex)
    If Val(cboYear.Tag) = 0 Then
        cboMonth.Enabled = False
        cboMonth.ListIndex = 0
    Else
        cboMonth.Enabled = True
        cboMonth.ListIndex = 0
    
        DTPFrom.Value = "1/1/" & Trim(Str(cboYear.Tag))
        DTPTo.Value = "12/31/" & Trim(Str(cboYear.Tag))
    End If
End If
End Sub

Private Sub cboMonth_Click()
If UserClick = True Then
    cboMonth.Tag = cboMonth.ItemData(cboMonth.ListIndex)
    If Val(cboMonth.Tag) = 0 Then
        ' Nothing
    Else
        Dim SetDay As String
        
        Select Case Val(cboMonth.Tag)
        Case 1, 3, 5, 7, 8, 10, 12
            SetDay = "31"
        Case 4, 6, 9, 11
            SetDay = "30"
        Case 2
            If (cboYear.ItemData(cboYear.ListIndex) Mod 4) = 0 Then
                SetDay = "29"
            Else
                SetDay = "28"
            End If
        End Select
    
        DTPFrom.Value = Trim(Str(cboMonth.Tag)) & "/1/" & Trim(Str(cboYear.ItemData(cboYear.ListIndex)))
        DTPTo.Value = Trim(Str(cboMonth.Tag)) & "/" & SetDay & "/" & Trim(Str(cboYear.ItemData(cboYear.ListIndex)))
    End If
End If
End Sub

Private Sub cmdExit_Click()
If UserClick = True Then Unload Me
End Sub


