Attribute VB_Name = "Global"
Option Explicit

'Programming:  Corey Michael Badeaux
'510 South Saint Patrick Street
'New Orleans, LA. 70119

'*********************************** GLOBAL VARIABLES ********************************************
Private SettingKey(1 To 255) As String  ' Settings from SETTINGS Table (Key).
Private SettingVal(1 To 255) As String  ' Settings from SETTINGS Table (Value).
Public DE As ADODB.Connection
Public UserClick As Boolean             'Makes sure Asychronous operations dont overlap and cause errors
Public FormPassedValue As Long          'Global Variable for FORM Parameters
Public FormPassedText As String         'Global Variable for FORM Parameters
Public ReportDir As String              'Report Directory Path
'Public EditChange As Boolean           'Marks When a Screen Edit has Occurred
'Public AddNew As Boolean               'Marks when a New mode is activated
'*************************************************************************************************
Public Declare Function CreateFieldDefFile Lib "p2smon.dll" (X As Object, ByVal fieldDefFilePath$, ByVal bOverWriteExistingFiles%) As Integer
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'*************************************************************************************************

Sub Main()
Dim SQLString As String

Select Case GetPC
Case "KINYODASX6"
    SQLString = "Provider=sqloledb;Data Source=KINYODASX6;Initial Catalog=EAST;User Id=sa;Password=nola0226"
    ReportDir = "D:\Data\Workspace\EastTree\EastTree\"
Case Else
    SQLString = "Provider=sqloledb;Data Source=kinyodas.servebbs.com,14033;Initial Catalog=EAST;User Id=sa;Password=nola0226"
    ReportDir = AppPath
End Select

Set DE = New ADODB.Connection
DE.CursorLocation = adUseClient
DE.Open SQLString

LoadSettings

FrmMain.Show
End Sub

Public Function GetPC() As String   'Return the name of the computer
Dim TempText As String
TempText = Space$(16)
If GetComputerName(TempText, Len(TempText)) <> 0 Then
   GetPC = TrimNull(TempText)
End If
End Function

Private Function TrimNull(Item As String)   'Trims the left side of the string up to the NULL
Dim TempInt As Integer
TempInt = InStr(Item, Chr$(0))
If TempInt Then
      TrimNull = Left$(Item, TempInt - 1)
Else: TrimNull = Item
End If
End Function

Public Sub Ending()
DE.Close: Set DE = Nothing
End
End Sub

Public Sub LoadForm(frm As Form, Optional LoadType As Integer, Optional PassValue As Long, Optional TabValue As Integer)    'Formats loaded forms to be centered.
If PassValue > 0 Then FormPassedValue = PassValue: DoEvents

frm.Move (FrmMain.Width / 2.2) - (frm.Width / 2.2), (FrmMain.Height / 2.5) - (frm.Height / 2.2)

Select Case LoadType
Case 0
    frm.Show
Case 1
    frm.Show vbModal
End Select

DoEvents
End Sub

Public Sub StartRoutine(frm As Form)
'Subroutine to A. Indicate Activity and B. Make actions aren't done twice.
UserClick = False
frm.MousePointer = 11
DoEvents
End Sub

Public Sub EndRoutine(frm As Form)
'Subroutine to A. Indicate Activity and B. Make actions aren't done twice.
frm.MousePointer = 0
UserClick = True
DoEvents
End Sub

'*************************************************************************************************

Public Function RSTemplate(TableName As String, IDName As String, ByVal IDValue As Long, Optional SqlCode As String) As ADODB.Recordset
Set RSTemplate = New ADODB.Recordset
If SqlCode = "" Then
    RSTemplate.Open ("SELECT * FROM " & TableName + " Where " & IDName & " = " & Trim(Str(IDValue))), DE, adOpenStatic, adLockOptimistic
Else
    RSTemplate.Open SqlCode, DE, adOpenStatic, adLockOptimistic
End If
End Function

Public Sub DeleteTemplate(TableName As String, IDName As String, IDValue As String, Optional IDName2 As String, Optional IDValue2 As String)
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
    'DELETE FROM [TABLE] WHERE [FIELD] = [VALUE]
    If Len(Trim(IDName2)) > 0 Then
        RS.Open ("DELETE FROM " & TableName & " WHERE " & IDName & " = " & IDValue & " And " & IDName2 & " = " & IDValue2), DE.SQL, adOpenStatic, adLockOptimistic
    Else
        RS.Open ("DELETE FROM " & TableName & " WHERE " & IDName & " = " & IDValue), DE, adOpenStatic, adLockOptimistic
    End If
Set RS = Nothing
DoEvents
End Sub

Public Function TransTemplate(SQLQuery As String, RecordName As String, Optional NoRecordMessage As String) As String
'Converts Mutliple Translations (Template)
Dim RS As ADODB.Recordset
Set RS = RSTemplate("", "", 0, SQLQuery)

If Not RS.EOF Then
    RS.MoveFirst
    TransTemplate = Trim(RS(RecordName).Value & "")
Else
    If Len(NoRecordMessage) > 0 Then
        TransTemplate = NoRecordMessage
    Else
        TransTemplate = "ERROR"
    End If
End If

CloseRS RS
Set RS = Nothing
End Function

Public Function FillGrid(SQLQuery As String, FlexGrid As MSHFlexGrid, Optional CaptionTotal As String, Optional StatusBox As TextBox, Optional FrameBox As Frame, Optional JumpOrUpdateButton As CommandButton, Optional AddButton As CommandButton) As Boolean
Dim Cnt As Long
Dim RS As ADODB.Recordset

    FlexGrid.MousePointer = 11

    Set RS = RSTemplate("", "", 0, Trim(SQLQuery))
    
    FlexGrid.Cols = RS.Fields.Count

    If IsRSOpen(RS) = True Then
        If Not RS.EOF Then
            Set FlexGrid.DataSource = RS
            FlexGrid.Visible = True
            If Not StatusBox Is Nothing Then StatusBox.Visible = False
            If Not FrameBox Is Nothing Then
                FrameBox.Enabled = True
                FrameBox.Caption = (CaptionTotal + " Totals: " + Trim(Str(RS.RecordCount)))
            End If
            If Not JumpOrUpdateButton Is Nothing Then JumpOrUpdateButton.Enabled = True
            If Not AddButton Is Nothing Then AddButton.Enabled = True
            FillGrid = True
        
            For Cnt = 0 To RS.Fields.Count - 1
                FlexGrid.TextMatrix(0, Cnt) = RS.Fields(Cnt).Name
                FlexGrid.ColWidth(Cnt) = 0
            Next
            RS.Close
            Set RS = Nothing
        Else
            FillGrid = False
        End If
    Else
        FillGrid = False
    End If
    
    If FillGrid = False Then
            FlexGrid.Clear
            FlexGrid.Visible = False
        If Not StatusBox Is Nothing Then StatusBox.Visible = True
        If Not FrameBox Is Nothing Then
            FrameBox.Enabled = False
            FrameBox.Caption = ""
        End If
        If Not JumpOrUpdateButton Is Nothing Then JumpOrUpdateButton.Enabled = False
        If Not AddButton Is Nothing Then AddButton.Enabled = False
    End If
    
    FlexGrid.MousePointer = 0
End Function

Public Function FillBox(ComNum As Object, TableName As String, RecordName As String, IDName As String, Optional SqlCode As String) As Long
'Fills up Various Combo or List Boxes
Dim Cnt As Long, Cnt2 As Long
Dim TempClick As Boolean
Dim RS As ADODB.Recordset

TempClick = UserClick
UserClick = False

If SqlCode = "" Then SqlCode = "SELECT " & RecordName & "," & IDName & " FROM " & TableName + " ORDER BY " & RecordName

Set RS = RSTemplate("", "", 0, SqlCode)

FillBox = RS.RecordCount
ComNum.Clear

If Not RS.EOF Then
    RS.MoveFirst
    For Cnt = 1 To FillBox
    If RS(RecordName) <> "" Then
        Cnt2 = Cnt2 + 1
        ComNum.AddItem Trim(RS(RecordName))
        ComNum.ItemData(Cnt2 - 1) = Val(Trim(RS(IDName).Value))
    End If
        RS.MoveNext
    Next Cnt
    If ComNum.ListCount > 0 Then ComNum.ListIndex = 0
End If

CloseRS RS
Set RS = Nothing

UserClick = TempClick
End Function

Public Sub CloseRS(adodbRS As ADODB.Recordset)
    If IsRSOpen(adodbRS) = True Then adodbRS.Close
End Sub

Public Function IsRSOpen(adodbRS As ADODB.Recordset) As Boolean
  If adodbRS.State And adStateOpen = adStateOpen Then
    IsRSOpen = True
  Else
    IsRSOpen = False
  End If
End Function

'*************************************************************************************************

Public Function FindItem(CBOCtrl As Object, Optional ListItem As String) As Integer
  Dim Cnt As Long
  Dim TempClick As Boolean
  
  TempClick = UserClick
  UserClick = False
  FindItem = -1
  
  For Cnt = 0 To CBOCtrl.ListCount - 1
    If (ListItem = CBOCtrl.List(Cnt) And Not ListItem = "") Or (ListItem = "" And Trim(CBOCtrl.Tag) = Trim(CBOCtrl.List(Cnt))) Then
      CBOCtrl.ListIndex = Cnt
      FindItem = Cnt
      Exit For
    End If
  Next Cnt
  
  UserClick = TempClick
End Function

Public Function FindID(CBOCtrl As Object, Optional ListID As Long) As Integer
  Dim Cnt As Long
  Dim TempClick As Boolean
  
  TempClick = UserClick
  UserClick = False
  FindID = -1
  
  For Cnt = 0 To CBOCtrl.ListCount - 1
    If (ListID = CBOCtrl.ItemData(Cnt) And Not ListID = 0) Or (ListID = 0 And Trim(CBOCtrl.Tag) = Trim(CBOCtrl.ItemData(Cnt))) Then
      CBOCtrl.ListIndex = Cnt
      FindID = Cnt
      Exit For
    End If
  Next Cnt
  
  UserClick = TempClick
End Function

Public Function SQLFix(ComText As String)
    'Fixes SQL searches by replacing all ' with '' to avoid error.
    SQLFix = Replace(ComText, "'", "''")
End Function

Public Function FindForm(ByVal FName As String) As Boolean    ' Return a form by name if it is loaded.
Dim Cnt As Long                                ' Search the loaded forms
    FindForm = False
    For Cnt = 0 To Forms.Count - 1
        If Forms(Cnt).Name = FName Then
            FindForm = True
            Exit For
        End If
    Next Cnt
End Function

Public Sub GridFix(ComNum As MSHFlexGrid)
    'Fixes Grids to highlight whole rows from the start.
    If ComNum.Col <> 0 And ComNum.ColSel <> (ComNum.Cols - 1) Then
        ComNum.Col = 0
        ComNum.ColSel = ComNum.Cols - 1
    End If
    ComNum.RowSel = ComNum.Row
End Sub

Public Function PickBox(PickType As Integer, Optional DefaultValue As String = "") As String
FormPassedValue = PickType
FormPassedText = DefaultValue

LoadForm FrmSearch, 1

PickBox = FormPassedValue
FormPassedValue = 0
End Function

Public Function AppPath() As String
If Right(App.Path, 1) <> "\" Then
    AppPath = App.Path + "\"
Else
    AppPath = App.Path
End If
End Function

Public Function PhoneFormat(ComNum As String, Optional FlagNum As Integer)
Dim Cnt As Long
Dim PN As String
For Cnt = 1 To Len(ComNum)
    If Asc(Mid(ComNum, Cnt, 1)) >= 48 And Asc(Mid(ComNum, Cnt, 1)) <= 57 Then
        PN = PN + Mid(ComNum, Cnt, 1)
    End If
Next Cnt

If Len(PN) = 10 Then            ' If it's not an American Phone Number.. leave it alone!
    If FlagNum = 0 Then         ' 0 - (xxx) xxx-xxxx Format
        PN = "(" + Left(PN + "0000000000", 10)
        PN = Left(PN, 4) + ") " + Mid(PN, 5, 3) + "-" + Right(PN, 4)
    ElseIf FlagNum = 1 Then     ' 1 - xxx-xxx-xxxx Format
        PN = "" + Left(PN + "0000000000", 10)
        PN = Left(PN, 3) + "-" + Mid(PN, 4, 3) + "-" + Right(PN, 4)
    End If
    PhoneFormat = Trim(PN)
Else
    PhoneFormat = ComNum
End If
End Function

Public Function NumbersOnly(KeyAscii As Integer)
    If KeyAscii > 26 And InStr("-01234567890.", Chr(KeyAscii)) = 0 Then
        NumbersOnly = 0
    Else
        NumbersOnly = KeyAscii
    End If
End Function

'*************************************************************************************************
Private Sub LoadSettings()
    Dim Cnt As Long
    Dim RS As ADODB.Recordset
        
        Set RS = RSTemplate("", "", 0, "Select Code as PKey, Value as PValue from Settings Order By Code")
    
    If RS.EOF = False Then
        RS.MoveFirst
        For Cnt = 1 To RS.RecordCount
            SettingKey(Cnt) = RS!PKey
            SettingVal(Cnt) = RS!PValue
            RS.MoveNext
        Next Cnt
    End If
    
    RS.Close
End Sub

Public Sub SettingSave()
    Dim Cnt As Long
    Dim RS As ADODB.Recordset
    
    Set RS = RSTemplate("", "", 0, "Select * From Settings Order By Code")
    
    If RS.EOF = False Then
        RS.MoveFirst
        For Cnt = 1 To RS.RecordCount
            RS!Value = SettingVal(Cnt)
            RS.MoveNext
        Next Cnt
    End If
    
    RS.Close
End Sub

Public Sub PutSetting(PKey As String, PValue As String)
    Dim Cnt As Long
    For Cnt = 1 To 255
        If UCase(Trim(PKey)) = UCase(Trim(SettingKey(Cnt))) Then
            SettingVal(Cnt) = PValue
            Exit Sub
        End If
    Next Cnt
End Sub

Public Function GetSetting(PKey As String) As String
    Dim Cnt As Long
    For Cnt = 1 To 255
        If UCase(Trim(PKey)) = UCase(Trim(SettingKey(Cnt))) Then
            GetSetting = UCase(Trim(SettingVal(Cnt)))
            Exit Function
        End If
    Next Cnt
    GetSetting = ""
End Function
'*************************************************************************************************

Public Function EditCheck(Optional CmdSaveButton, Optional CmdAddButton, Optional cmdCancelButton, Optional CmdDeleteButton, Optional CmdExitButton, Optional CmdDisableButton, Optional CmdDisableButton2, Optional CmdDisableButton3, Optional CmdDisableButton4, Optional CmdDisableButton5, Optional CmdDisableButton6, Optional CmdDisableButton7, Optional CmdDisableButton8, Optional CmdDisableButton9, Optional CmdDisableButton10) As Boolean
    'Simple SubRoutine only used to check if objects were edited or not by the user.
    If UserClick = True Then
        EditCheck = True
        If Not IsMissing(CmdSaveButton) Then CmdSaveButton.Enabled = True
        If Not IsMissing(CmdAddButton) Then CmdAddButton.Enabled = False
        If Not IsMissing(cmdCancelButton) Then cmdCancelButton.Enabled = True
        If Not IsMissing(CmdDeleteButton) Then CmdDeleteButton.Enabled = False
        If Not IsMissing(CmdExitButton) Then CmdExitButton.Enabled = False
        If Not IsMissing(CmdDisableButton) Then CmdDisableButton.Enabled = False
        If Not IsMissing(CmdDisableButton2) Then CmdDisableButton2.Enabled = False
        If Not IsMissing(CmdDisableButton3) Then CmdDisableButton3.Enabled = False
        If Not IsMissing(CmdDisableButton4) Then CmdDisableButton4.Enabled = False
        If Not IsMissing(CmdDisableButton5) Then CmdDisableButton5.Enabled = False
        If Not IsMissing(CmdDisableButton6) Then CmdDisableButton6.Enabled = False
        If Not IsMissing(CmdDisableButton7) Then CmdDisableButton7.Enabled = False
        If Not IsMissing(CmdDisableButton8) Then CmdDisableButton8.Enabled = False
        If Not IsMissing(CmdDisableButton9) Then CmdDisableButton9.Enabled = False
        If Not IsMissing(CmdDisableButton10) Then CmdDisableButton10.Enabled = False
    End If
End Function

Public Sub ClearFrame(vbForm As Form, vbFrame As Frame)
Dim VBObject As Object
Dim lhWnd As Long
On Error GoTo errFE
    For Each VBObject In vbForm
        lhWnd = VBObject.Container.hWnd
        If lhWnd <> 0 And lhWnd = vbFrame.hWnd Then
            If Not TypeOf VBObject Is Frame Then
                If TypeOf VBObject Is TextBox Or TypeOf VBObject Is ComboBox Or TypeOf VBObject Is DataCombo Then VBObject.Text = ""
                If TypeOf VBObject Is ListBox Or TypeOf VBObject Is ComboBox Then VBObject.ListIndex = 0
                If TypeOf VBObject Is CheckBox Then VBObject.Value = 0
                If TypeOf VBObject Is DTPicker Then VBObject.Value = ""
                DoEvents
            End If
        End If
    Next VBObject
    Exit Sub
errFE:
'    If Err = 438 Then lhWnd = 0:
    Resume Next
'    MsgBox "Error " & Err & vbNewLine & vbNewLine & Error, vbCritical, vbForm.Caption
    Err.Clear
End Sub

Public Sub LockTabs(TabObject As SSTab, TabSet As Long)    ' Locks all tabs except for the one selected
Dim Cnt As Long
TabObject.Tab = TabSet
For Cnt = 0 To (TabObject.Tabs - 1)
    If Cnt <> TabObject.Tab Then TabObject.TabEnabled(Cnt) = False
Next Cnt
End Sub
Public Sub UnLockTabs(TabObject As SSTab, Optional TabStart As Long = -1)    'Unlocks all tabs
Dim Cnt As Long
For Cnt = 0 To (TabObject.Tabs - 1)
    TabObject.TabEnabled(Cnt) = True
Next Cnt
If TabStart > -1 Then TabObject.Tab = TabStart
End Sub
