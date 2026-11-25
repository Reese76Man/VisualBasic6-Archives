VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Money Transfer"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   Icon            =   "frmTransfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   7395
   Begin VB.CommandButton cmdTransfer 
      Caption         =   "Transfer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6480
      Picture         =   "frmTransfer.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtFunds 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtAmount 
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.ComboBox cboItem 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmTransfer.frx":0884
      Left            =   720
      List            =   "frmTransfer.frx":0886
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1920
      Width           =   5175
   End
   Begin VB.TextBox txtNote 
      Height          =   1005
      Left            =   720
      MaxLength       =   2048
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2280
      Width           =   5175
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
      Left            =   6480
      Picture         =   "frmTransfer.frx":0888
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdPickEntT 
      Height          =   285
      Left            =   5880
      Picture         =   "frmTransfer.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   480
      Width           =   375
   End
   Begin VB.ComboBox cboEntityT 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   5055
   End
   Begin VB.CommandButton cmdPickEntF 
      Height          =   285
      Left            =   5880
      Picture         =   "frmTransfer.frx":0DCC
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.ComboBox cboEntityF 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
   Begin MSComCtl2.DTPicker DTPWorkDate 
      Height          =   315
      Left            =   720
      TabIndex        =   3
      Top             =   1560
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
      Format          =   108593155
      CurrentDate     =   27908
   End
   Begin VB.Label Label2 
      Caption         =   "Available Funds:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Line Line3 
      X1              =   6360
      X2              =   120
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      X1              =   6360
      X2              =   6360
      Y1              =   120
      Y2              =   3240
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6360
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label15 
      Caption         =   "Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label21 
      Caption         =   "Amount:"
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label22 
      Caption         =   "Item:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label17 
      Caption         =   "Notes:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "To:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label14 
      Caption         =   "From:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTransfer_Click()
If UserClick = True And Val(txtAmount.Text) > 0 And cboEntityF.ItemData(cboEntityF.ListIndex) <> cboEntityT.ItemData(cboEntityT.ListIndex) Then
        
        If MsgBox("Funds are about to be transferred.  Proceed?", vbYesNo, "Fund Transfer") = vbYes Then
            Dim RS As ADODB.Recordset
            Set RS = RSTemplate("WorkDetail", "WorkDetailID", 0)
            RS.AddNew
                    
            RS!WorkOrderID = 0
            RS!WorkDate = DTPWorkDate.Value
            RS!WorkItemID = Val(cboItem.ItemData(cboItem.ListIndex))
            RS!FromEntityID = Val(cboEntityF.ItemData(cboEntityF.ListIndex))
            RS!EntityID = Val(cboEntityT.ItemData(cboEntityT.ListIndex))
            RS!Description = Trim(txtNote.Text)
            RS!Quantity = 1
            RS!Amount = Val(txtAmount.Text)
            
            RS.Update
            CloseRS RS
            
            MsgBox "Transaction Processed!", vbInformation
            
            Set RS = RSTemplate("", "", 0, "Exec EastTreeUpdater 0")
            Set RS = Nothing
            
            ClearForm
            cboEntityF_Click
        End If
End If
End Sub

Private Sub Form_Activate()
cboEntityF.SetFocus
End Sub

Private Sub Form_Load()
cboEntityF.Tag = FillBox(cboEntityF, "", "ListItem", "ListID", "Exec EntityList 4")
cboEntityT.Tag = FillBox(cboEntityT, "", "ListItem", "ListID", "Exec EntityList 5")
cboItem.Tag = FillBox(cboItem, "", "ListItem", "ListID", "Exec ItemList 2, 4")

ClearForm
End Sub

Private Sub ClearForm()
DTPWorkDate.Value = Date$
cboEntityF.ListIndex = 0
cboEntityT.ListIndex = 0
cboItem.ListIndex = 0
txtAmount.Text = ""
txtNote.Text = ""
End Sub

Private Sub cboEntityF_Click()
'If UserClick = True Then
    StartRoutine Me
        Dim RS As ADODB.Recordset
        Set RS = RSTemplate("", "", 0, "Exec EntityMain " & cboEntityF.ItemData(cboEntityF.ListIndex))
        
        If Not RS.EOF Then
            txtFunds.Text = Val("" & RS!Paid)
        End If
                
        CloseRS RS
        Set RS = Nothing
    EndRoutine Me
'End If
End Sub


Private Sub cmdExit_Click()
If UserClick = True Then Unload Me
End Sub

Private Sub cmdPickEntF_Click()
If UserClick = True Then
    cmdPickEntF.Tag = PickBox(8)
    
    If Val(cmdPickEntF.Tag) > 0 Then cboEntityF.Tag = FindID(cboEntityF, Val(cmdPickEntF.Tag))
End If
End Sub
Private Sub cmdPickEntT_Click()
If UserClick = True Then
    cmdPickEntT.Tag = PickBox(8)
    
    If Val(cmdPickEntT.Tag) > 0 Then cboEntityT.Tag = FindID(cboEntityT, Val(cmdPickEntT.Tag))
End If
End Sub
