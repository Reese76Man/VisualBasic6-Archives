VERSION 5.00
Begin VB.Form frmForms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forms"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   9885
   Begin VB.ComboBox cboForm 
      Height          =   315
      ItemData        =   "frmForms.frx":0000
      Left            =   1200
      List            =   "frmForms.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   5895
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
      Height          =   1215
      Left            =   8640
      Picture         =   "frmForms.frx":001C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Preview / Print Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7320
      Picture         =   "frmForms.frx":045E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Form:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   7200
      X2              =   7200
      Y1              =   1320
      Y2              =   0
   End
End
Attribute VB_Name = "frmForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
StartRoutine Me
    cboForm.ListIndex = 0
EndRoutine Me
End Sub

Private Sub cmdPrint_Click()
If UserClick = True Then
    StartRoutine Me
        Dim RS As ADODB.Recordset
        Set RS = RSTemplate("", "", 0, "Exec FormReport " & cboForm.ItemData(cboForm.ListIndex))
        
        frmReport.ViewReport "WorkOrderForm", RS, , False
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

