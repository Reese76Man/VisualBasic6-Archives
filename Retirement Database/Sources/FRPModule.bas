Attribute VB_Name = "FRPModule"
Option Explicit
'*********************************** API ROUTINES & VARIABLES ***********************************
Private Type udtBrowse
    lngHwnd        As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "Kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As udtBrowse) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'-----------------------------------------------------------
'Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function CreateFieldDefFile Lib "p2smon.dll" (X As Object, ByVal fieldDefFilePath$, ByVal bOverWriteExistingFiles%) As Integer
'************************************************************************************************

'*********************************** GLOBAL VARIABLES ********************************************
Public SQLDate As String            'Current Date from SQL Server (or otherwise)
'Public NTUser As String             'Currently Logged in NT User  (or otherwise)
Public PCName As String             'Currently PC Name
'Public EmployeeID As Long           'Logged in EmployeeID of User
'Public EmployeeFullName As String   'Logged in Employee FullName
'Public SecurityName As String       'Security Settings Description
'Public SecurityID As Long           'SecurityID for Employee
Public MyPath As String             'Local Application Path
Public SavePath As String           'Path to which files are saved to.
'----------------------------------  BACKGROUND VARIABLES ----------------------------------------
Public UserClick As Boolean         'Makes sure Asychronous operations dont overlap and cause errors
Public ShowForm As Boolean          'Used when forms are terminated for display purposes.
Public FormPassedValue As String    'Global Variable for FORM Parameters
Public FormCurrentValue As String   'Global Variable for FORM Parameters
Public FormPickedValue As String    'Global Variable for Pick FORM Parameters
Public EditChange As Boolean        'Checks to see if object was edited.
Public EditChange2 As Boolean        'Checks to see if object was edited.
Public EditChange3 As Boolean        'Checks to see if object was edited.
Public EditChange4 As Boolean        'Checks to see if object was edited.
Public EditChange5 As Boolean        'Checks to see if object was edited.
Public EditChange6 As Boolean        'Checks to see if object was edited.
Public AddNew As Boolean            'Checks to see if ADD was in progress.
Public AddNew2 As Boolean            'Checks to see if ADD was in progress.
Public AddNew3 As Boolean            'Checks to see if ADD was in progress.
Public AddNew4 As Boolean            'Checks to see if ADD was in progress.
Public AddNew5 As Boolean            'Checks to see if ADD was in progress.
Public AddNew6 As Boolean            'Checks to see if ADD was in progress.
Public FormHistory(1 To 500) As Form      'Keeps track of previous forms
Public FormValue(1 To 500) As String      'Keeps track of form parameters
Public FormCount As Integer               'Keeps track of form count
Public AddItem As Long              'Long Integer of Choices for ADDITEM form & Feedback ID (1 - Customer)
Public AddItem2 As Long             '.. Choice only
Public AddItem3 As Single           '.. Choice only
Public AddItem4 As Long             '.. Choice only
Private Security(1 To 100) As String * 1 'R - Read Only | F - Full Access | O - No Access
Public ScanOutput As Integer        'Form in which Scanned Data is sent.
'*************************************************************************************************


Public Sub LoadForm(frm As Form, Optional LoadType As Integer)    'Formats loaded forms to be centered.
    If LoadType <> 0 Then AddItem = LoadType
    
    frm.Move (MDIMain.Width / 2.2) - (frm.Width / 2.2), (MDIMain.Height / 2.5) - (frm.Height / 2.2)
    
    If LoadType <> 0 Then
        frm.Show vbModal
    Else
        frm.Show
    End If
    DoEvents
End Sub

Public Function FileExists(ComVar As String) As Boolean
If Not (ComVar = "" Or Right(ComVar, 1) = "\") Then
    FileExists = (Dir(ComVar) <> "")
Else
    FileExists = False
End If
End Function

Public Function PathExists(ComVar As String) As Boolean
If Dir$(ComVar, vbDirectory) = "" Then
    PathExists = False
Else
    PathExists = True
End If
End Function

Public Sub SavePathCheck(ComFrm As Form)
Do Until PathExists(SavePath) = True
If PathExists(SavePath) = False Then
    SavePath = BrowseFolder(ComFrm.hwnd, "Select Save Folder")
End If
Loop
End Sub

Public Sub GetSQLDate()             'Get Global Information (Date from SQL, User from NT)
    'DE.SQLDate
    '    SQLDate = Trim(DE.rsSQLDate!SQLDate)
    SQLDate = Date$
    'DE.rsSQLDate.Close
End Sub                             'End Routine to get SQL Server Date

Public Sub GetUserData()            'Routine to call function to get User Login Name
    'NTUser = GetUser
    PCName = GetPC
    'EmployeeID = GetUserID(NTUser)
    'SecurityID = GetSecurityID(EmployeeID)
    'EmployeeFullName = EmployeeID2Name(EmployeeID)
    MyPath = PathFix(App.Path)
    SavePath = MyPath
'    If SecurityCheck(SecurityID) = False Then
'        MsgBox "Error Loading Security Settings!", vbCritical
'        ExitProgram
'    End If
End Sub                             'End User Login Name Function Call

Public Sub OpenWindow(frm As Form, ComVar As String)  'Simple Subroutine to Open SubWindows and Shut down properly
    EditChange = False
'    EditChange2 = False
    AddNew = False
    ShowForm = True
    FormValue(FormCount + 1) = ComVar
    Set FormHistory(FormCount + 1) = frm
    FormPassedValue = FormValue(FormCount + 1)
    Unload FormHistory(FormCount)
    FormCount = FormCount + 1
    DoEvents
    LoadForm FormHistory(FormCount)
    ShowForm = False
    DoEvents
End Sub

Public Sub CloseWindow()            'Simple Subroutine to Close SubWindows and Shut down properly
    FormPassedValue = ""
    ShowForm = True
    If FormCount < 2 Then FormCount = 1: Exit Sub
    Unload FormHistory(FormCount)
    FormCount = FormCount - 1
    DoEvents
    FormPassedValue = FormValue(FormCount)
    LoadForm FormHistory(FormCount)
    Set FormHistory(FormCount + 1) = Nothing
    ShowForm = False
    DoEvents
End Sub

Public Sub Xout(ComVar As String)   'Routine that makes it where when a form is terminated, it either reverts
                                    ' back to the previous form, nothing, or the main menu
    If ShowForm = False Then
        MDIMain.MNUMainMenu.Enabled = True
            If MsgBox("Do you want to go back to Main Menu?", vbYesNo, "Exit Screen") = vbYes Then
                MDIMain.MNUMainMenu.Enabled = False
                FormCount = 1
                Set FormHistory(FormCount) = frmMainMenu
                LoadForm frmMainMenu
            End If
        Exit Sub
    Else
        FormValue(FormCount) = Trim(ComVar)
    End If
End Sub

Public Sub StartProgram()
    LoadRecords
    GetSQLDate
    GetUserData
    'DE.Data.Open
End Sub

Public Sub LoadRecords()
    DE.EntryFullNames
    'DE.CourseNumbers
End Sub

Public Sub UnLoadRecords()
    DE.rsEntryFullNames.Close
'    DE.rsCourseNumbers.Close
End Sub

Public Sub ExitProgram()            'Routines to do when program fully unloads.
                                    'Close all OPEN Recordsets and Connections.
    UnLoadRecords
    DE.Data.Close
    End
End Sub

Public Sub EditCheck(Optional CmdSaveButton, Optional CmdAddButton, Optional cmdCancelButton)
    'Simple SubRoutine only used to check if objects were edited or not by the user.
    If UserClick = True Then
        EditChange = True
        If Not IsMissing(CmdSaveButton) Then CmdSaveButton.Enabled = True
        If Not IsMissing(CmdAddButton) Then CmdAddButton.Enabled = False
        If Not IsMissing(cmdCancelButton) Then cmdCancelButton.Enabled = True
    End If
End Sub

Public Sub EditCheck2(Optional CmdSaveButton, Optional CmdAddButton, Optional cmdCancelButton)
    'Simple SubRoutine only used to check if objects were edited or not by the user.
    If UserClick = True Then
        EditChange2 = True
        If Not IsMissing(CmdSaveButton) Then CmdSaveButton.Enabled = True
        If Not IsMissing(CmdAddButton) Then CmdAddButton.Enabled = False
        If Not IsMissing(cmdCancelButton) Then cmdCancelButton.Enabled = True
    End If
End Sub
Public Sub EditCheck3(Optional CmdSaveButton, Optional CmdAddButton, Optional cmdCancelButton)
    'Simple SubRoutine only used to check if objects were edited or not by the user.
    If UserClick = True Then
        EditChange3 = True
        If Not IsMissing(CmdSaveButton) Then CmdSaveButton.Enabled = True
        If Not IsMissing(CmdAddButton) Then CmdAddButton.Enabled = False
        If Not IsMissing(cmdCancelButton) Then cmdCancelButton.Enabled = True
    End If
End Sub
Public Sub EditCheck4(Optional CmdSaveButton, Optional CmdAddButton, Optional cmdCancelButton)
    'Simple SubRoutine only used to check if objects were edited or not by the user.
    If UserClick = True Then
        EditChange4 = True
        If Not IsMissing(CmdSaveButton) Then CmdSaveButton.Enabled = True
        If Not IsMissing(CmdAddButton) Then CmdAddButton.Enabled = False
        If Not IsMissing(cmdCancelButton) Then cmdCancelButton.Enabled = True
    End If
End Sub

Public Sub StartRoutine(frm As Form) 'Subroutine to A. Indicate Activity and B. Make actions aren't done twice.
    UserClick = False
    frm.MousePointer = 11
    DoEvents
End Sub

Public Sub EndRoutine(frm As Form)  'Subroutine to A. Indicate Activity and B. Make actions aren't done twice.
    frm.MousePointer = 0
    UserClick = True
    DoEvents
End Sub

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

Public Function FindItem(CBOCtrl As ComboBox, ListItem As String) As Integer
  Dim CNT As Long
  Dim boolUC As Boolean
  
  boolUC = UserClick
  UserClick = False
  FindItem = -1
  
  For CNT = 0 To CBOCtrl.ListCount - 1
    If ListItem = CBOCtrl.List(CNT) Then
      CBOCtrl.ListIndex = CNT
      FindItem = CNT
      Exit For
    End If
  Next CNT
  
  UserClick = boolUC
End Function

Public Function FillTemplate(comnum As ComboBox, TableName As String, RecordName As String, IDName As String, Optional SqlCode As String) As String
'Fills up Various Combo Boxes (GLOBAL TEMPLATE I)
Dim CNT, CNT2 As Long
Dim TempClick As Boolean
Dim RS As Recordset
TempClick = UserClick
UserClick = False
Set RS = New Recordset

If SqlCode = "" Then
    RS.Open ("SELECT " & RecordName & "," & IDName & " FROM " & TableName + " ORDER BY " & RecordName), DE.Data, adOpenStatic, adLockReadOnly
Else
    RS.Open SqlCode, DE.Data, adOpenStatic, adLockReadOnly
End If

If Not RS.EOF Then
    RS.MoveFirst
    For CNT = 1 To RS.RecordCount
    If RS(RecordName) <> "" Then
        CNT2 = CNT2 + 1
        comnum.AddItem Trim(RS(RecordName))
        comnum.ItemData(CNT2 - 1) = Val(Trim(RS(IDName).Value))
    End If
        RS.MoveNext
    Next CNT
    If comnum.ListCount > 0 Then comnum.ListIndex = 0
Else
    FillTemplate = "ERROR"
End If
RS.Close
Set RS = Nothing
UserClick = TempClick
End Function

Public Function SaveTemplate(TableName As String, IDName As String, ByVal IDValue As Long, Optional SqlCode As String) As Recordset
'Save Template for Edit-Add-Saves
Set SaveTemplate = New Recordset
If SqlCode = "" Then
    SaveTemplate.Open ("SELECT * FROM " & TableName + " Where " & IDName & " = " & Trim(Str(IDValue))), DE.Data, adOpenStatic, adLockOptimistic
Else
    SaveTemplate.Open SqlCode, DE.Data, adOpenStatic, adLockReadOnly
End If
End Function


Public Function NumTrim(Number As Double, Places As Integer) As Double ' Simple Yet Effective Number Trimming
NumTrim = Number * (10 ^ Places)
NumTrim = Round(NumTrim)
NumTrim = NumTrim / (10 ^ Places)
End Function


Public Sub DeleteTemplate(TableName As String, IDName As String, IDValue As String)
'Deletes records in a table based on a certain criteria
Dim RS As Recordset
Set RS = New Recordset
    'DELETE FROM [TABLE] WHERE [FIELD] = [VALUE]
    RS.Open ("DELETE FROM " & TableName & " WHERE " & IDName & " = " & IDValue), DE.Data, adOpenStatic, adLockOptimistic
    'MsgBox ("DELETE FROM " & TableName & " WHERE " & IDName & " = " & IDValue)  'Test Before Deleting!
Set RS = Nothing
DoEvents
End Sub

Public Function Exists(TableName As String, Parameter1 As String, Value1 As String, Optional Parameter2 As String, Optional Value2 As String, Optional Parameter3 As String, Optional Value3 As String) As Boolean
'Finds out if a Record Exists
Dim CNT As Long
Dim RS As Recordset
Set RS = New Recordset

If Parameter2 = "" And Parameter3 = "" Then
        RS.Open ("SELECT " & Parameter1 & " FROM " & TableName & " Where " & Parameter1 & " = " & Value1), DE.Data, adOpenStatic, adLockReadOnly
ElseIf Parameter3 = "" Then
        RS.Open ("SELECT " & Parameter1 & "," & Parameter2 & " FROM " & TableName & " Where " & Parameter1 & " = " & Value1 & " and " & Parameter2 & " = " & Value2), DE.Data, adOpenStatic, adLockReadOnly
Else
        RS.Open ("SELECT " & Parameter1 & "," & Parameter2 & "," & Parameter3 & " FROM " & TableName & " Where " & Parameter1 & " = " & Value1 & " and " & Parameter2 & " = " & Value2 & " and " & Parameter3 & " = " & Value3), DE.Data, adOpenStatic, adLockReadOnly
End If

If Not RS.EOF Then
    Exists = True
Else
    Exists = False
End If

RS.Close
Set RS = Nothing
End Function

Public Function PhoneNumberFormat(comnum As String, Optional FlagNum As Integer)
Dim CNT As Long
Dim PN As String
For CNT = 1 To Len(comnum)
    If Asc(Mid(comnum, CNT, 1)) >= 48 And Asc(Mid(comnum, CNT, 1)) <= 57 Then
        PN = PN + Mid(comnum, CNT, 1)
    End If
Next CNT

If Len(PN) = 10 Then            ' If it's not an American Phone Number.. leave it alone!
    If FlagNum = 0 Then         ' 0 - (xxx) xxx-xxxx Format
        PN = "(" + Left(PN + "0000000000", 10)
        PN = Left(PN, 4) + ") " + Mid(PN, 5, 3) + "-" + Right(PN, 4)
    ElseIf FlagNum = 1 Then     ' 1 - xxx-xxx-xxxx Format
        PN = "" + Left(PN + "0000000000", 10)
        PN = Left(PN, 3) + "-" + Mid(PN, 4, 3) + "-" + Right(PN, 4)
    End If
    PhoneNumberFormat = Trim(PN)
Else
    PhoneNumberFormat = comnum
End If
End Function

'********************************** CONVERSION QUERIES *******************************
Public Function GetUserID(ComVar As String) As Long 'Converts NTLoginName to EmployeeID
    GetUserID = Val(TransTemplate("select EmployeeID from Employee WHERE NTLogin = '" & Trim(UCase(ComVar)) & "'", "EmployeeID", 1))
End Function
Public Function GetSecurityID(ComVar As Long) As Long 'Converts EmployeeID to SecurityID
    GetSecurityID = Val(TransTemplate("select SecurityID from Employee WHERE EmployeeID = '" & Trim(UCase(ComVar)) & "'", "SecurityID", 1))
End Function

Public Function SecurityCheck(comnum As Long, Optional ComVar As String) As Boolean
'If ComVar = "" Then
'    DE.SecurityCheck ComNum
'    If Not DE.rsSecurityCheck.EOF Then
'        Security(1) = UCase(Trim(DE.rsSecurityCheck!Applicant))     'Applicant SECURITY - 1
'        SecurityName = Trim(DE.rsSecurityCheck!SecurityDescription)
'        SecurityCheck = True
'    Else
'        SecurityCheck = False
'    End If
'    DE.rsSecurityCheck.Close
'Else
'    If Security(ComNum) = UCase(Trim(ComVar)) Then
        SecurityCheck = True
'    Else
'        SecurityCheck = False
'    End If
'End If
End Function

Public Function BrowseFolder(ByVal lngHwnd As Long, ByVal sMessage As String) As String
    On Error GoTo ExitHandler
    Dim intNull As Integer
    Dim lngIDList As Long, lngResult As Long
    Dim strPath As String
    Dim udtBI As udtBrowse
    'Setting API properties
    With udtBI
        .lngHwnd = lngHwnd
        .lpszTitle = lstrcat(sMessage, "")
        .ulFlags = 1
    End With
    'Displaying the browse folder dialog box...
    lngIDList = SHBrowseForFolder(udtBI)
    If lngIDList <> 0 Then
        'Create string of nulls so it will fill in with the path
        strPath = String(260, 0)
        'Retrieves the path selected, places in the null
         'character filled string
        lngResult = SHGetPathFromIDList(lngIDList, strPath)
        'Release memory
        Call CoTaskMemFree(lngIDList)
        'Find the first instance of a null character,
         'to get just the path
        intNull = InStr(strPath, vbNullChar)
        'Greater than 0 means path exists...
        If intNull > 0 Then
            'Set the value
            strPath = Left(strPath, intNull - 1)
        End If
    End If
    'Return the path
    BrowseFolder = PathFix(strPath)
    Exit Function
ExitHandler:
    BrowseFolder = Empty
End Function

Public Function BrowseFile(BrowserFlag As String, CDB As Object, Optional Title As String, Optional FFilter As String) As String
CDB.CancelError = True
    On Error GoTo ErrHandler
    Select Case BrowserFlag
        Case "C"            '
            CDB.Flags = cdlCCFullOpen
            CDB.ShowColor
        Case "F"            ' Font Dialog
            CDB.Flags = cdlCFBoth + cdlCFTTOnly
            CDB.FontName = "Times New Roman"
            CDB.FontSize = 20
            CDB.ShowFont
        Case "H"            ' Help Dialog
            CDB.HelpFile = "VB5.HLP"
            CDB.HelpCommand = cdlHelpContents
            CDB.ShowHelp
        Case "O"            ' Open File Dialog
            CDB.DialogTitle = Trim(Title)
            If Trim(Title) = "" Then CDB.DialogTitle = "File Open"
            If Len(FFilter) > 0 Then
                CDB.Filter = Trim(FFilter)
            Else
                CDB.Filter = "All Images|*.jpg;*.gif;*.bmp"
            End If
            CDB.Flags = cdlOFNAllowMultiselect + cdlOFNExplorer + cdlOFNLongNames + cdlOFNFileMustExist
            'CDB.FileName = "*.Jpg"
            'CDB.DefaultExt = "Jpg"
            CDB.ShowOpen
            BrowseFile = Trim(CDB.FileName)
        Case "P"            ' Print Dialog
            CDB.Flags = cdlPDHidePrintToFile
            CDB.ShowPrinter
        Case "S"            ' Save Dialog
            CDB.DialogTitle = Trim(Title)
            CDB.Filter = "Visual Basic Files (*.Vbp)|*.vbp|All Files (*.*)|*.*"
            CDB.Flags = cdlOFNOverwritePrompt + cdlOFNHelpButton + cdlOFNPathMustExist
            'CDB.Flags = 2066
            CDB.ShowSave
        Case Else
            MsgBox "Whoops!"
    End Select
    On Error GoTo 0 'Reset error handling
    'User didn't press cancel.  Take an action with dialog value.
    Exit Function
ErrHandler:
    If Err = cdlCancel Then
        BrowseFile = ""
    End If
End Function

Public Function FirstRecord(SourceTable As String, ReturnField As String, Order As String) As String
FirstRecord = TransTemplate("Select " & ReturnField & " From " & SourceTable & " ORDER BY " & Order, ReturnField)
End Function

Public Sub LockTabs(TabObject As SSTAB)     ' Locks all tabs except for the one selected
Dim Tmp As Long
For Tmp = 0 To (TabObject.Tabs - 1)
    If Tmp <> TabObject.Tab Then TabObject.TabEnabled(Tmp) = False
Next Tmp
End Sub
Public Sub UnLockTabs(TabObject As SSTAB)   'Unlocks all tabs
Dim Tmp As Long
For Tmp = 0 To (TabObject.Tabs - 1)
    TabObject.TabEnabled(Tmp) = True
Next Tmp
End Sub
Public Sub GridFix(GridObj As MSHFlexGrid)
    'Fixes Grids to highlight whole rows from the start.
    If GridObj.Col <> 0 And GridObj.ColSel <> (GridObj.Cols - 1) Then
        GridObj.Col = 0
        GridObj.ColSel = GridObj.Cols - 1
    End If
    GridObj.RowSel = GridObj.Row
End Sub

Public Function DateCheck(DateData As String) As String ' Makes sure a date is valid.
DateData = Trim(DateData)
If IsDate(DateData) = True Then
    DateCheck = DateData
Else
    DateCheck = ""
End If
End Function

Public Sub ClearFrame(vbForm As Form, vbFrame As Frame)
Dim VBObject As Object
Dim lhWnd As Long
On Error GoTo errFE
    For Each VBObject In vbForm
        lhWnd = VBObject.Container.hwnd
        If lhWnd <> 0 And lhWnd = vbFrame.hwnd Then
            If Not TypeOf VBObject Is Frame Then
'                MsgBox VBObject.Text
                VBObject.Text = ""
                VBObject.ListIndex = 0
                VBObject.Value = 0
                DoEvents
            End If
        End If
    Next VBObject
    Exit Sub
errFE:
    If Err = 438 Then lhWnd = 0: Resume Next
'    MsgBox "Error " & Err & vbNewLine & vbNewLine & Error, vbCritical, vbForm.Caption
    Err.Clear
End Sub

Public Function FindForm(ByVal Fname As String) As Boolean    ' Return a form by name if it is loaded.
Dim CNT As Long                                ' Search the loaded forms
    FindForm = False
    For CNT = 0 To Forms.Count - 1
        If Forms(CNT).Name = Fname Then
            FindForm = True
            Exit For
        End If
    Next CNT
End Function

Public Function YesNoValue(comnum As String) As Byte    ' Function to convert for CheckBoxes
If Trim(UCase(comnum)) = "Y" Then
    YesNoValue = 1
Else
    YesNoValue = 0
End If
End Function

Public Function NameFormat(comnum As Variant) As String
If Len(comnum) > 1 Then
    NameFormat = Trim(UCase(Left("" & comnum, 1)) & LCase(Right("" & comnum, Len(comnum) - 1)))
Else
    NameFormat = Trim(UCase("" & comnum))
End If
End Function

Public Sub MSCOMInit(MSC As MSCommLib.MSComm)
On Error Resume Next
MSC.CommPort = 1
MSC.Settings = "9600,N,8,1"   ' 9600 baud, no parity, 8 data, and 1 stop bit.
MSC.InputLen = 0              ' Read entire buffer
MSC.PortOpen = True           ' Opens the port
MSC.RThreshold = 1            ' Call **_OnComm for each character
End Sub

Public Sub MSCOMUnInit(MSC As MSCommLib.MSComm)
If MSC.PortOpen = True Then MSC.PortOpen = False
End Sub

Public Function MSCOMRecieve(MSC As MSCommLib.MSComm, Optional ComFlag As Integer) As String
Dim TextRecieved As String
Dim CNT As Long
  Select Case MSC.CommEvent
    Case comEvReceive   ' Received RThreshold # of chars.
        TextRecieved = MSC.Input
        For CNT = 1 To Len(TextRecieved)
            If Asc(Mid(TextRecieved, CNT, 1)) > 31 Then MSC.Tag = MSC.Tag & Mid(TextRecieved, CNT, 1)
            If Asc(Mid(TextRecieved, CNT, 1)) = 10 Then
                MSCOMRecieve = MSC.Tag
                MSC.Tag = ""
            End If
        Next CNT
  End Select
End Function

Public Function LoadImage(ComImage As Image, ImagePath As String, PlaceHoldShape As Shape) As Boolean
On Error GoTo ExitHandler
Dim ImageRatio As Double   ' Loads a Picture and Resizes it into an Image Control
If FileExists(ImagePath) = True Then
    Set ComImage.Picture = LoadPicture(ImagePath)
    If ComImage.Picture.Height > ComImage.Picture.Width Then
        ImageRatio = ComImage.Picture.Width / ComImage.Picture.Height
        ComImage.Height = PlaceHoldShape.Height
        ComImage.Width = PlaceHoldShape.Width * ImageRatio
    ElseIf ComImage.Picture.Width > ComImage.Picture.Height Then
        ImageRatio = ComImage.Picture.Height / ComImage.Picture.Width
        ComImage.Height = PlaceHoldShape.Height * ImageRatio
        ComImage.Width = PlaceHoldShape.Width
    End If
    ComImage.Stretch = True
    ComImage.Tag = Trim(ImagePath)
    LoadImage = True
Else
    LoadImage = False
End If
Exit Function

ExitHandler:
LoadImage = False
End Function

Public Sub MailTo(ComEmail As String)
    Dim Web As Object
    Set Web = CreateObject("InternetExplorer.Application")
    Web.Navigate "mailto:" & Trim(ComEmail) & "?Subject=<Enter Subject>"
    Set Web = Nothing
End Sub

Public Function ValidEmail(ByVal ComEmail As String) As Boolean
        Dim i As Integer
        Dim atplace As Integer
        Dim dotplace As Integer
        Dim atcount As Integer
        Dim atcountboo As Boolean
        Dim dotcount As Integer

        For i = 1 To Len(ComEmail)
            If Mid(ComEmail, i, 1) = "@" Then 'Checks where the location of the @ symbol is and if there is more than 1
                atcount = atcount + 1
                atplace = i
                If atcount > 1 Then ValidEmail = False
            End If
            If Mid(ComEmail, i, 1) = "." Then
                dotcount = dotcount + 1
                dotplace = i
            End If
            If InStr("abcdefghijklmnopqrstuvwxyz1234567890@.-_", LCase(Mid(ComEmail, i, 1))) > 0 Then 'checks for illegal charaters
            Else
                ValidEmail = False
            End If
        Next

        If dotplace < atplace Or Len(ComEmail) = dotplace Then ValidEmail = False

        If atcount <> 1 Then 'checks if there are invalid or no @ symbols
            atcountboo = False
        Else
            atcountboo = True
        End If
        
        If Len(ComEmail) > 0 Then
            If Mid(ComEmail, 1, 1) = "@" Or Mid(ComEmail, Len(ComEmail), 1) = "@" Then ValidEmail = False
        Else
            ValidEmail = False
        End If

        If Mid(ComEmail, atplace + 1, 1) = "." Then ValidEmail = False

        If dotcount >= 1 And atcountboo = True Then 'checks if there is at least 1 . and an @ exists in the right place
            ValidEmail = True
        Else
            ValidEmail = False
        End If
End Function

Public Sub SetCommonTableFields(TableName As String, CommonField As String, ByVal CommonFieldParameter As Long, FieldToSet As String, SetFieldValue As Boolean, Optional SingleFieldToSet As Boolean, Optional SingleCommonField As String, Optional ByVal SingleCommonFieldParameter As Long)
Dim RS As Recordset
Dim CNT As Long
Set RS = SaveTemplate(TableName, CommonField, CommonFieldParameter)
    
If Not RS.EOF Then
    For CNT = 1 To RS.RecordCount
        RS.Move (CNT - 1), adBookmarkFirst
        If Len(SingleCommonField) > 0 Then
            If RS(SingleCommonField).Value = SingleCommonFieldParameter Then
                RS(FieldToSet).Value = SingleFieldToSet
            Else
                RS(FieldToSet).Value = SetFieldValue
            End If
        Else
            RS(FieldToSet).Value = SetFieldValue
        End If
        RS.Update
    Next CNT
End If
   
CloseRS RS
Set RS = Nothing
End Sub

'*********************************** PRIVATE FUNCTIONS & SUBROUTINES *****************
Private Function GetPC() As String     'Return the name of the computer
   Dim Tmp As String
   Tmp = Space$(16)
   If GetComputerName(Tmp, Len(Tmp)) <> 0 Then
      GetPC = TrimNull(Tmp)
   End If
End Function
Private Function PathFix(ComVar As String) As String
If Not Right(ComVar, 1) = "\" Then PathFix = ComVar + "\"
End Function

'Private Function GetUser() As String
  'return the name of the user
'   Dim Tmp As String
'   Tmp = Space$(256)
'   If GetUserName(Tmp, Len(Tmp)) Then
'      GetUser = TrimNull(Tmp)
'   End If
'End Function

Private Function TrimNull(item As String)   'Trims the left side of the string up to the NULL
   Dim pos As Integer
   pos = InStr(item, Chr$(0))
   If pos Then
         TrimNull = UCase(Left$(item, pos - 1))
   Else: TrimNull = item
   End If
End Function
Private Function TransTemplate(SQLQuery As String, RecordName As String, Optional ConnectionFlag As Byte) As String  'Converts Mutliple Translations (Template)
    Dim RS As Recordset
    Set RS = New Recordset
    RS.MaxRecords = 1
    If ConnectionFlag = 0 Then
        RS.Open SQLQuery, DE.Data, adOpenStatic, adLockReadOnly
    ElseIf ConnectionFlag = 1 Then
    '    RS.Open SQLQuery, DE.GLOBAL, adOpenStatic, adLockReadOnly
    End If
    If Not RS.EOF Then
        RS.MoveFirst
        TransTemplate = Trim(RS(RecordName).Value & "")
    Else
        TransTemplate = "ERROR"
    End If
    RS.Close
    Set RS = Nothing
End Function
'**************************************************************************************
