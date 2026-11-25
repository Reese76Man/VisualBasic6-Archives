VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm FrmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "East Tree LLC"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10725
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8265
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu MNUFile 
      Caption         =   "File"
      Begin VB.Menu MNUMainMenu 
         Caption         =   "Main Menu"
      End
      Begin VB.Menu MNUExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MNUInvoice 
      Caption         =   "Invoices"
   End
   Begin VB.Menu MNUAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
'
End Sub

Private Sub MDIForm_Activate()
LoadForm FrmMenu
End Sub

Private Sub MNUAbout_Click()
LoadForm frmAbout
End Sub

Private Sub MNUInvoice_Click()
LoadForm frmInvoice
End Sub

Private Sub MNUMainMenu_Click()
LoadForm FrmMenu
End Sub

Private Sub MNUExit_Click()
Unload Me
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
Ending
End Sub


