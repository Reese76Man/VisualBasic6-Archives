VERSION 5.00
Begin VB.Form frmInvoice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Report"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   Icon            =   "frmInvoice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   8325
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Preview / Print Invoice"
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
      Left            =   7440
      Picture         =   "frmInvoice.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
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
      Left            =   7440
      Picture         =   "frmInvoice.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   855
   End
   Begin VB.ComboBox cboWork 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   5895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   7320
      X2              =   7320
      Y1              =   3120
      Y2              =   120
   End
   Begin VB.Label Label1 
      Caption         =   "Work Orders:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
StartRoutine Me
    cboWork.Tag = FillBox(cboWork, "", "WorkDesc", "WorkOrderID", "Select WorkOrderID, (Right(('000000' + Convert(VarChar,WorkOrderID)),7) + ': ' + Convert(VarChar,OpenDate,101) + ' - ' + [Description]) As WorkDesc From WorkOrder Where CloseDate Is Null Order By OpenDate Desc, [Description]")
EndRoutine Me
End Sub

Private Sub cmdPrint_Click()
If UserClick = True Then
    Dim BillOut As Integer
    'If MsgBox("Is this Invoice ready to bill out?", vbYesNo, "Finalize Invoice") = vbYes Then
    '    BillOut = 1
    'End If
        
    StartRoutine Me
        Dim RS As ADODB.Recordset
        Set RS = RSTemplate("", "", 0, "Exec InvoiceReport " & cboWork.ItemData(cboWork.ListIndex) & "," & Trim(Str(BillOut)))
        
        frmReport.ViewReport "Invoice", RS, , False
        LoadForm frmReport, 1
        
        CloseRS RS
        Set RS = Nothing
    EndRoutine Me
End If
End Sub

' *************************************************************************************************

Private Sub cmdExit_Click()
If UserClick = True Then Unload Me
End Sub
