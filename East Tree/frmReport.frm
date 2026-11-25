VERSION 5.00
Object = "{FB992564-9055-42B5-B433-FEA84CEA93C4}#11.0#0"; "CRViewer.dll"
Begin VB.Form frmReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Preview"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8265
   WindowState     =   2  'Maximized
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
      Height          =   440
      Left            =   7560
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer CRViewer 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      _cx             =   14420
      _cy             =   11245
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1033
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CRAppl As CRAXDRT.Application
Dim Report As CRAXDRT.Report

Public Sub ViewReport(rptName As String, RS As ADODB.Recordset, Optional ExportOptions As Byte, Optional MakeTTX As Boolean)
    'ReportDir = GetSetting("PATH") & "Reports\"

   If MakeTTX = True Then
        Call CreateFieldDefFile(RS, ReportDir & "EAST.ttx", True)
        Exit Sub
   End If
  
  Set CRAppl = New CRAXDRT.Application
  Set Report = CRAppl.OpenReport(ReportDir & rptName & ".rpt")
  
  Report.Database.Tables(1).SetDataSource RS
  Report.ReadRecords
  
'  If ExportOptions > 0 Then
'    Report.ExportOptions.DestinationType = crEDTDiskFile
'    Select Case ExportOptions
'    Case 1
'        Report.ExportOptions.FormatType = crEFTExactRichText
'        Report.ExportOptions.DiskFileName = SavePath & rptName & ".RTF"
'    Case 2
'        Report.ExportOptions.FormatType = crEFTWordForWindows
'        Report.ExportOptions.DiskFileName = SavePath & rptName & ".DOC"
'    Case 3
'        Report.ExportOptions.FormatType = crEFTExcel97
'        Report.ExportOptions.DiskFileName = SavePath & rptName & ".XLS"
'    Case 4
'        Report.ExportOptions.FormatType = crEFTText
'        Report.ExportOptions.DiskFileName = SavePath & rptName & ".TXT"
'    End Select
'    Report.ExportOptions.ExcelExportAllPages = True
'    Report.Export False
'    Unload Me
'  Else
    CRViewer.ReportSource = Report
    CRViewer.ViewReport
'  End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub CRViewer_PrintButtonClicked(UseDefault As Boolean)
  Dim hWnd As Long
  UseDefault = False
  Report.PrinterSetup (hWnd)
  Report.PrintOut True
End Sub

Private Sub Form_Resize()
  CRViewer.Top = 0
  CRViewer.Left = 0
  CRViewer.Height = ScaleHeight
  CRViewer.Width = ScaleWidth
  cmdExit.Left = ScaleWidth - cmdExit.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set CRAppl = Nothing
  Set Report = Nothing
End Sub

