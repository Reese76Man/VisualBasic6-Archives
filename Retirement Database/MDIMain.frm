VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm MDIMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "FRP"
   ClientHeight    =   11550
   ClientLeft      =   -105
   ClientTop       =   465
   ClientWidth     =   13230
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "DBE User Status Bar."
      Top             =   11295
      Width           =   13230
      _ExtentX        =   23336
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      Enabled         =   0   'False
   End
   Begin MSCommLib.MSComm MSCOM 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComDlg.CommonDialog CDBox 
      Left            =   0
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu MNUFile 
      Caption         =   "&File"
      Begin VB.Menu MNUMainMenu 
         Caption         =   "&Main Menu"
         Shortcut        =   ^M
      End
      Begin VB.Menu MNUExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu MNUReports 
      Caption         =   "&Reports"
      Begin VB.Menu MNURetireeReports 
         Caption         =   "&Retiree Reports"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu MNUAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Load()
    StartProgram
    
    'MDIMain.StatusBar.SimpleText = "UserID:" & Str(EmployeeID) & "        NTUser: " & NTUser & "        PCID: " & PCName & "        Date: " & SQLDate & "        Security Access: " & SecurityName
    MDIMain.StatusBar.SimpleText = "PCID: " & PCName & "        Date: " & SQLDate
    
'    If EmployeeID <= 0 Then
'        MsgBox (NTUser & " is not in our database.  Contact Administrator.")
'        End
'    End If
    
'    MSCOMInit MSCOM
    
    MNUMainMenu_Click
End Sub

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And MNUReports.Enabled = True Then
        PopupMenu MNUReports
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Makes it where any unloaded MAIN FORMS will unload without asking to go to Main Menu!
    ShowForm = True
End Sub

Private Sub MDIForm_Resize()
    ' Will resize current form in case user decides to resize MDI form.
    If FormCount > 0 Then LoadForm FormHistory(FormCount)
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'MSCOMUnInit MSCOM
    ExitProgram
    End
End Sub

Private Sub MNUAbout_Click()
    MNUAbout.Enabled = False
    LoadForm frmAbout
End Sub

Private Sub MNUExit_Click()
    Unload Me
End Sub

Private Sub MNUMainMenu_Click()
    ShowForm = False
    MNUMainMenu.Enabled = False
    FormCount = 1
    Set FormHistory(FormCount) = frmMainMenu
    LoadForm frmMainMenu
End Sub

Private Sub MNURetireeReports_Click()
    MNUReports.Enabled = False
    LoadForm frmEntryReports
End Sub

Private Sub MSCOM_OnComm()
    Dim ScannedText As String
    ScannedText = MSCOMRecieve(MSCOM)
    If ScannedText > "" Then
        If FindForm("frmAbout") = True Then frmAbout.txtBarCode.Text = ScannedText
    End If
End Sub
