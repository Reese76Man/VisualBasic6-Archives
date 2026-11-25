VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmReport 
   Caption         =   "Report Preview"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7110
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
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
      Height          =   430
      Left            =   6000
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      lastProp        =   500
      _cx             =   12515
      _cy             =   7435
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
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

Public Sub ViewReport(rptName As String, adodbRS As ADODB.Recordset, Optional ExportOptions As Byte)
'  Call CreateFieldDefFile(adodbRS, MyPath & "test.ttx", True)
'  Exit Sub
  
  Set CRAppl = New CRAXDRT.Application
  'Set Report = CRAppl.OpenReport(DefaultPath & rptName & ".rpt")
  Set Report = CRAppl.OpenReport(MyPath & rptName & ".rpt")
  Report.Database.Tables(1).SetDataSource adodbRS
  Report.ReadRecords
  
'  If ExportOptions > 0 Then
'    Report.ExportOptions.DestinationType = crEDTDiskFile
'    If ExportOptions = 1 Then
'        Report.ExportOptions.FormatType = crEFTExactRichText
'        Report.ExportOptions.DiskFileName = SavePath & rptName & ".RTF"
'    ElseIf ExportOptions = 2 Then
'        Report.ExportOptions.FormatType = crEFTWordForWindows
'        Report.ExportOptions.DiskFileName = SavePath & rptName & ".DOC"
'    ElseIf ExportOptions = 3 Then
'        Report.ExportOptions.FormatType = crEFTExcel97
'        Report.ExportOptions.DiskFileName = SavePath & rptName & ".XLS"
'    End If
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
  Dim hwnd As Long
  UseDefault = False
  Report.PrinterSetup (hwnd)
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
