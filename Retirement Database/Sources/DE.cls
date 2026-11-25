VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Data As ADODB.Connection

Public rsEntryFullNames As ADODB.Recordset
Public rsEntryMain As ADODB.Recordset
Public rsMembershipMain As ADODB.Recordset
Public rsAddressMain As ADODB.Recordset
Public rsEmailMain As ADODB.Recordset
Public rsEntrySearch As ADODB.Recordset
Public rsDuesMain As ADODB.Recordset
Public rsEntryReport1 As ADODB.Recordset
Public rsMembershipGroup As ADODB.Recordset
Public rsEntryReport2 As ADODB.Recordset
Public rsChapterMain As ADODB.Recordset
Public rsOccupationMain As ADODB.Recordset

Private Sub Class_Initialize()
    Set Data = New ADODB.Connection
    Data.Provider = "Microsoft.Jet.OLEDB.4.0"

    Set rsEntryFullNames = New ADODB.Recordset
    Set rsEntryMain = New ADODB.Recordset
    Set rsMembershipMain = New ADODB.Recordset
    Set rsAddressMain = New ADODB.Recordset
    Set rsEmailMain = New ADODB.Recordset
    Set rsEntrySearch = New ADODB.Recordset
    Set rsDuesMain = New ADODB.Recordset
    Set rsEntryReport1 = New ADODB.Recordset
    Set rsMembershipGroup = New ADODB.Recordset
    Set rsEntryReport2 = New ADODB.Recordset
    Set rsChapterMain = New ADODB.Recordset
    Set rsOccupationMain = New ADODB.Recordset

    DataEnvironment_Initialize
End Sub

Public Sub DataEnvironment_Initialize()
    If Data.State = adStateOpen Then Data.Close
    Data.Provider = "Microsoft.Jet.OLEDB.4.0"
    Data.Properties("Data Source").Value = MyPath & "data.mdb"
End Sub

Private Sub EnsureConnection()
    If Data.State <> adStateOpen Then
        If Len(Data.Properties("Data Source").Value & "") = 0 Then
            Data.Properties("Data Source").Value = MyPath & "data.mdb"
        End If
        Data.Open
    End If
End Sub

Private Sub OpenParameterized(ByVal rs As ADODB.Recordset, ByVal sql As String, ParamArray params() As Variant)
    EnsureConnection
    If rs.State = adStateOpen Then rs.Close

    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = Data
    cmd.CommandType = adCmdText
    cmd.CommandText = sql

    Dim i As Long
    For i = LBound(params) To UBound(params)
        Dim p As ADODB.Parameter
        If VarType(params(i)) = vbBoolean Then
            Set p = cmd.CreateParameter("P" & i, adBoolean, adParamInput, , params(i))
        Else
            Set p = cmd.CreateParameter("P" & i, adVarWChar, adParamInput, 510, CStr(params(i)))
        End If
        cmd.Parameters.Append p
    Next i

    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic, adLockOptimistic
End Sub

Public Sub EntryFullNames()
    EnsureConnection
    If rsEntryFullNames.State = adStateOpen Then rsEntryFullNames.Close
    rsEntryFullNames.CursorLocation = adUseClient
    rsEntryFullNames.Open "SELECT EntryID, LastName + ', ' + FirstName + ' ' + MiddleName AS FullName FROM Entry ORDER BY LastName, FirstName, MiddleName", Data, adOpenStatic, adLockOptimistic
End Sub

Public Sub EntryMain(ByVal EntryNumber As Variant)
    Dim sql As String
    sql = "SELECT Entry.*, Address.*, Email.*, LastName + ', ' + FirstName + ' ' + MiddleName AS FullName " & _
          "FROM (Entry LEFT JOIN Address ON Entry.EntryID = Address.EntryID) LEFT JOIN Email ON Entry.EntryID = Email.EntryID " & _
          "WHERE Entry.EntryID = ? " & _
          "ORDER BY Address.PreferredAddress, Email.PreferredAddress;"
    OpenParameterized rsEntryMain, sql, EntryNumber
End Sub

Public Sub MembershipMain(ByVal EntryNum As Variant)
    Dim sql As String
    sql = "SELECT MembershipGroupAssociation.MembershipGroupAssociationID, MembershipGroupAssociation.MembershipGroupID, MembershipGroup.Description " & _
          "FROM MembershipGroup, MembershipGroupAssociation " & _
          "WHERE MembershipGroup.MembershipGroupID = MembershipGroupAssociation.MembershipGroupID " & _
          "AND MembershipGroupAssociation.EntryID = ? " & _
          "ORDER BY MembershipGroup.Description"
    OpenParameterized rsMembershipMain, sql, EntryNum
End Sub

Public Sub AddressMain(ByVal EntryNum As Variant)
    Dim sql As String
    sql = "SELECT Address.AddressID, Address.Address1, Address.Address2, Address.City, Address.State, Address.Zipcode, Address.Country, Address.Phone, Address.BusinessAddress, Address.PreferredAddress " & _
          "FROM ADDRESS " & _
          "WHERE Address.EntryID = ? " & _
          "ORDER BY Address.PreferredAddress ASC"
    OpenParameterized rsAddressMain, sql, EntryNum
End Sub

Public Sub EmailMain(ByVal EntryNum As Variant)
    Dim sql As String
    sql = "SELECT Email.EmailID, Email.EmailAddress, Email.PreferredAddress " & _
          "FROM Email " & _
          "WHERE Email.EntryID = ? " & _
          "ORDER BY Email.PreferredAddress ASC"
    OpenParameterized rsEmailMain, sql, EntryNum
End Sub

Public Sub EntrySearch(ByVal EntryText As Variant)
    Dim sql As String
    sql = "SELECT EntryID, LastName + ', ' + FirstName + ' ' + MiddleName AS FullName, Active, NationalChapterNumber, NationalChapterDate " & _
          "FROM Entry " & _
          "WHERE (LastName + ', ' + FirstName + ' ' + MiddleName + ' ') LIKE '%' + ? + '%'"
    OpenParameterized rsEntrySearch, sql, EntryText
End Sub

Public Sub DuesMain(ByVal EntryNum As Variant)
    Dim sql As String
    sql = "SELECT Dues.* FROM Dues WHERE Dues.EntryID = ? ORDER BY PaymentDate DESC"
    OpenParameterized rsDuesMain, sql, EntryNum
End Sub


Private Function ToBool(ByVal v As Variant) As Boolean
    Dim s As String
    s = LCase$(Trim$(CStr(v & "")))
    Select Case s
        Case "true", "-1", "1", "yes"
            ToBool = True
        Case Else
            ToBool = False
    End Select
End Function

Public Sub EntryReport1(ByVal ActiveStatus As Variant, ByVal ActiveStatus2 As Variant, ByVal DecStatus As Variant, ByVal DecStatus2 As Variant)
    Dim sql As String
    sql = "SELECT Entry.*, Address.*, Email.*, FirstName + ' ' + MiddleName + ' ' + LastName AS FullName " & _
          "FROM (Entry LEFT JOIN Address ON Entry.EntryID = Address.EntryID) LEFT JOIN Email ON Entry.EntryID = Email.EntryID " & _
          "WHERE ((Address.PreferredAddress = True OR Address.PreferredAddress Is Null) AND (Email.PreferredAddress = True OR Email.PreferredAddress Is Null)) " & _
          "AND (Entry.Active = ? OR NOT Entry.Active = ?) AND (Entry.Deceased = ? OR NOT Entry.Deceased = ?)"
    OpenParameterized rsEntryReport1, sql, ToBool(ActiveStatus), ToBool(ActiveStatus2), ToBool(DecStatus), ToBool(DecStatus2)
End Sub


Public Sub MembershipGroup()
    EnsureConnection
    If rsMembershipGroup.State = adStateOpen Then rsMembershipGroup.Close
    rsMembershipGroup.CursorLocation = adUseClient
    rsMembershipGroup.Open "SELECT * FROM MembershipGroup", Data, adOpenStatic, adLockOptimistic
End Sub

Public Sub EntryReport2(ByVal ComGroupID As Variant)
    Dim sql As String
    sql = "SELECT Entry.*, Address.*, Email.*, FirstName + ' ' + MiddleName + ' ' + LastName AS FullName, MembershipGroup.Description " & _
          "FROM (((Entry LEFT JOIN Address ON Entry.EntryID = Address.EntryID) LEFT JOIN Email ON Entry.EntryID = Email.EntryID) " & _
          "INNER JOIN MembershipGroupAssociation ON Entry.EntryID = MembershipGroupAssociation.EntryID) " & _
          "INNER JOIN MembershipGroup ON MembershipGroupAssociation.MembershipGroupID = MembershipGroup.MembershipGroupID " & _
          "WHERE (((Address.PreferredAddress) = True OR (Address.PreferredAddress) Is Null) " & _
          "AND ((Email.PreferredAddress) = True OR (Email.PreferredAddress) Is Null) " & _
          "AND MembershipGroupAssociation.MembershipGroupID = ?)"
    OpenParameterized rsEntryReport2, sql, ComGroupID
End Sub

Public Sub ChapterMain(ByVal EntryNum As Variant)
    Dim sql As String
    sql = "SELECT LocalChapter.LocalChapterID, LocalChapter.LocalChapterDate, LocalChapter.LocalChapterNumber " & _
          "FROM LocalChapter " & _
          "WHERE LocalChapter.EntryID = ? " & _
          "ORDER BY LocalChapter.LocalChapterDate"
    OpenParameterized rsChapterMain, sql, EntryNum
End Sub

Public Sub OccupationMain(ByVal EntryNum As Variant)
    Dim sql As String
    sql = "SELECT Occupation.OccupationID, Occupation.Occupation, Occupation.Classification " & _
          "FROM Occupation " & _
          "WHERE Occupation.EntryID = ? " & _
          "ORDER BY Occupation.Occupation"
    OpenParameterized rsOccupationMain, sql, EntryNum
End Sub