VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRSSubjectNumbering"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000 All Rights Reserved
'   File:       clsRSSubjectNumbering.cls
'   Author:     Nicky Johns November 2000
'   Purpose:    Subject Numbering class for use during Registration
'               in MACRO Registration/Randomisation module
'------------------------------------------------------------------------------------------------'
'
'------------------------------------------------------------------------------------------------'
'   Revisions:
'   NCJ 16-20 Nov Initial development
'   NCJ 5 Dec 00 - Tidy up as result of white box test by TA
'   NCJ/TA 29 Jan 03 - Ensure correct database stuff for MACRO 3.0
'   NCJ 26 Mar 03 - Added time zone offset in StoreUniquenessChecks
'------------------------------------------------------------------------------------------------'

Option Explicit

Private msTrialName As String
Private msSite As String
Private mlPersonId As Long
Private msPrefixValue As String
Private msSuffixValue As String

Private mlStartNumber As Long
Private mnNumberWidth As Integer
Private mbUsePrefix As Boolean
Private mbUseSuffix As Boolean
Private msSubjectIdentifier As String

' Store a result message
Private msResultMsg As String
Private mnResultCode As Integer

Private msChecksTimeStamp As String
' Store uniqueness checks for this patient
Private mcolSubjectChecks As clsRSSubjectChecks

'------------------------------------------------------------------------------------'
Public Sub SetDatabase(ByVal sDatabaseName As String)
'------------------------------------------------------------------------------------'
' Set the database to connect to
' NCJ/TA 29 Jan 03 - Ensure correct database stuff for MACRO 3.0
'------------------------------------------------------------------------------------'
Dim sSecCon As String
    
    Set goUser = New MACROUser

    sSecCon = InitializeSecurityADODBConnection
    
    Set goUser = SilentUserLogin(sSecCon, sDatabaseName, "")

    InitializeMacroADODBConnection

End Sub


'---------------------------------------------------------------------
Private Function SilentUserLogin(sSecCon As String, ByVal sDatabase As String, sDefaultHTMLLocation As String) As MACROUser
'---------------------------------------------------------------------
Dim oUser As MACROUser
Dim bLoginSucceeded As Boolean
Dim sMessage As String

    Set oUser = New MACROUser
    bLoginSucceeded = oUser.SilentLogin(sSecCon, sDatabase, sDefaultHTMLLocation, sMessage)
    
    If bLoginSucceeded Then
        Set SilentUserLogin = oUser
    Else
        Set SilentUserLogin = Nothing
    End If
    

End Function

'----------------------------------------------------------------------------------------'
Public Sub Initialise(ByVal sTrialName As String, _
                        ByVal sSite As String, _
                        ByVal lPersonId As Long, _
                        ByVal sPrefix As String, _
                        ByVal sSuffix As String, _
                        ByVal nUsePrefix As Integer, _
                        ByVal nUseSuffix As Integer, _
                        ByVal lStartNumber As Long, _
                        ByVal nNumberWidth As Integer)
'----------------------------------------------------------------------------------------'
' Initialise numbering class for this subject
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
    
    ' Store values in class variables
    msTrialName = sTrialName
    msSite = sSite
    mlPersonId = lPersonId
    msPrefixValue = sPrefix
    msSuffixValue = sSuffix
    mbUsePrefix = (nUsePrefix = 1)
    mbUseSuffix = (nUseSuffix = 1)
    mlStartNumber = lStartNumber
    mnNumberWidth = nNumberWidth
    
    ' Initialise local values for new object
    msChecksTimeStamp = ""
    msSubjectIdentifier = ""
    msResultMsg = ""
    mnResultCode = 0
    Set mcolSubjectChecks = New clsRSSubjectChecks

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "Initialise", "clsRSSubjectNumbering")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'----------------------------------------------------------------------------------------'
Public Sub AddUniquenessChecks(ByVal sUCheckString As String)
'----------------------------------------------------------------------------------------'
' sUCheckString should contain uniqueness check codes and values of form
'   Code1|Value1|Code2|Value2| ... etc.
' May be empty string, but if not, expect a terminating |
'----------------------------------------------------------------------------------------'
Dim lStrPos As Long
Dim lSepPos As Long
Dim sCheckCode As String
Dim sCheckValue As String

Const sSep = "|"

    If sUCheckString = "" Then Exit Sub
    
    lStrPos = 1       ' Start at the beginning
    Do
        ' Look for next separator
        lSepPos = InStr(lStrPos, sUCheckString, sSep, vbBinaryCompare)
        If lSepPos > 0 Then
            ' Read the check code
            sCheckCode = Mid(sUCheckString, lStrPos, lSepPos - lStrPos)
            lStrPos = lSepPos + 1   ' Move past the separator
            lSepPos = InStr(lStrPos, sUCheckString, sSep, vbBinaryCompare)
            If lSepPos > 0 Then
                ' We expect the value to follow!
                sCheckValue = Mid(sUCheckString, lStrPos, lSepPos - lStrPos)
                Call Me.AddUniquenessCheckValue(sCheckCode, sCheckValue)
                lStrPos = lSepPos + 1   ' Move past the separator
            End If
        Else
            ' No more separators found
            Exit Do
        End If
    Loop


End Sub

'----------------------------------------------------------------------------------------'
Private Function IsUnique() As Boolean
'----------------------------------------------------------------------------------------'
' See if the current patient is unique according to patient check values
' Assume the collection of check values for this patient is filled in
' If unique, the current patient's checks are added to the DB (otherwise they're NOT stored)
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsChecks As ADODB.Recordset
Dim sCurSite As String
Dim lCurPersonId As Long
Dim nMatches As Integer
Dim oSubjectCheck As clsRSSubjectCheck
Dim bMisMatch As Boolean
Dim bIsUnique As Boolean

    On Error GoTo ErrHandler
    
    ' Initialise result
    IsUnique = True
    
    ' See if there is anything to check
    ' The mcolSubjectChecks should contain the uniqueness checks
    If mcolSubjectChecks.Count = 0 Then Exit Function
    
    ' See if this person has already been done
    sSQL = "SELECT * FROM RSUniquenessCheck " _
            & " WHERE ClinicalTrialName = '" & msTrialName & "'" _
            & " AND TrialSite = '" & msSite & "'" _
            & " AND PersonId = " & mlPersonId
    Set rsChecks = New ADODB.Recordset
    rsChecks.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    ' If person has already been checked in then leave
    If rsChecks.RecordCount > 0 Then
        rsChecks.Close
        Set rsChecks = Nothing
        Exit Function
    End If
    
    ' Get all the existing patient check values
    sSQL = "SELECT TrialSite, PersonId, CheckCode, CheckValue " _
            & " FROM RSUniquenessCheck " _
            & " WHERE ClinicalTrialName = '" & msTrialName & "'" _
            & " ORDER BY TrialSite, PersonId "
    Set rsChecks = New ADODB.Recordset
    rsChecks.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    ' Initialise current site and current person etc.
    bIsUnique = True
    sCurSite = ""
    lCurPersonId = -1
    nMatches = 0
    bMisMatch = False
    Do While Not rsChecks.EOF
        If (rsChecks!TrialSite <> sCurSite) Or (rsChecks!PersonId <> lCurPersonId) Then
            ' New person's record - check previous person first
            If nMatches = mcolSubjectChecks.Count Then
                ' The previous person matched on ALL codes!
                ' Set result and jump out of loop
                bIsUnique = False
                Exit Do
            End If
            ' Reset current site/person
            sCurSite = rsChecks!TrialSite
            lCurPersonId = rsChecks!PersonId
            nMatches = 0
            bMisMatch = False
        End If
        ' Only need to check record if we haven't yet found a mismatch
        If Not bMisMatch Then
            ' Get the current value for this check code (if any)
            Set oSubjectCheck = mcolSubjectChecks.Item(rsChecks!CheckCode)
            If Not oSubjectCheck Is Nothing Then
                If LCase(oSubjectCheck.CheckValue) = LCase(rsChecks!CheckValue) Then
                    ' Add this to number of matches
                    ' NB We need this AS WELL AS bMisMatch
                    nMatches = nMatches + 1
                Else
                    ' Mark that we found a mismatch
                    bMisMatch = True
                End If
            End If
        End If
        ' Move to next record
        rsChecks.MoveNext
    Loop
    
    ' We need to check the last person's records if we hit EOF
    ' (because we won't have gone through the above loop)
    If rsChecks.EOF And nMatches = mcolSubjectChecks.Count Then
        bIsUnique = False
    End If
    
    rsChecks.Close
    Set rsChecks = Nothing
    
    ' Store the values if uniqueness succeeds
    If bIsUnique Then
        Call StoreUniquenessChecks
    End If
    
    IsUnique = bIsUnique

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "IsUnique", "clsRSSubjectNumbering")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Function

'----------------------------------------------------------------------------------------'
Public Sub AddUniquenessCheckValue(ByVal sCheckCode As String, ByVal sCheckValue As String)
'----------------------------------------------------------------------------------------'
' Add a new Uniqueness check value
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
    
    If sCheckValue = "" Then Exit Sub
    
    ' Set the timestamp for the first check value to arrive
    ' and initialise patient check collection
    If msChecksTimeStamp = "" Then
        msChecksTimeStamp = SQLStandardNow
        Set mcolSubjectChecks = New clsRSSubjectChecks
    End If
    
    ' Add it to collection of patient check values for this patient
    Call mcolSubjectChecks.Add(sCheckCode, sCheckValue)

Exit Sub
ErrHandler:
  Err.Raise Err.Number, , Err.Description & "|clsRSSubjectNumbering.AddUniquenessCheckValue"
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub StoreUniquenessChecks()
'----------------------------------------------------------------------------------------'
' Store all uniqueness checks in the RSUniquenessChecks table
' by iterating through our mcolSubjectChecks collection
' Use current class property values for Trial, Site, Person and ChecksTimeStamp
' NCJ 26 Mar 03 - Added time zone value
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim oSubjectCheck As clsRSSubjectCheck
Dim oTimeZone As TimeZone

    On Error GoTo ErrHandler
    
    Set oTimeZone = New TimeZone
    
    For Each oSubjectCheck In mcolSubjectChecks
        With oSubjectCheck
            sSQL = "INSERT INTO RSUniquenessCheck " _
                & "(ClinicalTrialName, TrialSite, PersonId, " _
                & " CheckCode, CheckValue, CheckTimeStamp, " _
                & " CheckTimeStamp_TZ)" _
                & " VALUES ( '" & msTrialName & "', '" & msSite & "', " & mlPersonId _
                & ", '" & .CheckCode & "', '" & ReplaceQuotes(.CheckValue) & "', " & msChecksTimeStamp _
                & ", " & oTimeZone.TimezoneOffset & " )"
        
            MacroADODBConnection.Execute sSQL
        End With
    Next
    
    Set oTimeZone = Nothing

Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "StoreUniquenessChecks", "clsRSSubjectNumbering")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'----------------------------------------------------------------------------------------'
Public Property Get PrefixValue() As String
'----------------------------------------------------------------------------------------'
' Return value of Prefix (set in Initialise)
'----------------------------------------------------------------------------------------'

    PrefixValue = msPrefixValue
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get SuffixValue() As String
'----------------------------------------------------------------------------------------'
' Return value of Suffix (set in Initialise)
'----------------------------------------------------------------------------------------'

    SuffixValue = msSuffixValue

End Property

'----------------------------------------------------------------------------------------'
Public Property Get SubjectIdentifier() As String
'----------------------------------------------------------------------------------------'
' Return subject identifier for this person (returns empty string if cannot be calculated)
' Calculate it if local copy is empty, otherwise return private value
'----------------------------------------------------------------------------------------'

    If msSubjectIdentifier = "" Then
        Call GenerateIdentifier
    End If
    SubjectIdentifier = msSubjectIdentifier
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get ResultMessage() As String
'----------------------------------------------------------------------------------------'
' Return the result message which says what went wrong
' Returns empty string if there were no problems
'----------------------------------------------------------------------------------------'

    ResultMessage = msResultMsg
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get ResultCode() As Integer
'----------------------------------------------------------------------------------------'
' Return the result code which indicates what happened
' Returns 0 if there were no problems
'----------------------------------------------------------------------------------------'

    ResultCode = mnResultCode

End Property

'----------------------------------------------------------------------------------------'
Private Function GenerateIdentifier() As String
'----------------------------------------------------------------------------------------'
' Generate new identifier from current class property values
' Check if already registered or not unique and set return codes accordingly
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsIDS As ADODB.Recordset
Dim lNumber As Long
Dim rsSubjectId As ADODB.Recordset

    On Error GoTo ErrHandler
    
    ' Check we have a trial name, a site and a person
    If msTrialName = "" Or msSite = "" Or mlPersonId = 0 Then
        SetReturnMessage eRegResult.RegMissingInfo
        Exit Function
    End If
    
    ' See if there's already an identifier for this person
    sSQL = "SELECT * FROM RSSubjectIdentifier " _
            & " WHERE " & GetSQLStringEquals("ClinicalTrialName", msTrialName) _
            & " AND " & GetSQLStringEquals("TrialSite", msSite) _
            & " AND PersonId = " & mlPersonId
    Set rsSubjectId = New ADODB.Recordset
    rsSubjectId.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    If rsSubjectId.RecordCount > 0 Then
        ' This person already has an identifier
        msSubjectIdentifier = rsSubjectId!SubjectIdentifier
        rsSubjectId.Close
        Set rsSubjectId = Nothing
        SetReturnMessage eRegResult.RegAlreadyRegistered
        Exit Function
    End If
    
    ' For a new person need to check uniqueness
    If Not IsUnique Then
        SetReturnMessage eRegResult.RegNotUnique
        rsSubjectId.Close
        Set rsSubjectId = Nothing
        Exit Function
    End If
    
    ' Otherwise create a new record and leave recordset open so we can update it later
    rsSubjectId.AddNew
    rsSubjectId!ClinicalTrialName = msTrialName
    rsSubjectId!TrialSite = msSite
    rsSubjectId!PersonId = mlPersonId
    
    
    ' Now look for next number in the sequence
    sSQL = "SELECT * FROM RSNextNumber " _
        & " WHERE " & GetSQLStringEquals("ClinicalTrialName", msTrialName) _
        & " AND "
    If mbUsePrefix And msPrefixValue > "" Then
        sSQL = sSQL & GetSQLStringEquals("Prefix", msPrefixValue)
    Else
        'TA 4/12/00: changed '=' to 'is' for oracle
        sSQL = sSQL & "Prefix is NULL "
    End If
    sSQL = sSQL & " AND "
    If mbUseSuffix And msSuffixValue > "" Then
        sSQL = sSQL & GetSQLStringEquals("Suffix", msSuffixValue)
    Else
        'TA 4/12/00: changed '=' to 'is' for oracle
        sSQL = sSQL & "Suffix is NULL "
    End If

    Set rsIDS = New ADODB.Recordset
    rsIDS.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    If rsIDS.RecordCount > 0 Then
        ' There is a number there already
        lNumber = rsIDS!NextNumber
    Else
        lNumber = mlStartNumber
        ' Create a new record
        rsIDS.AddNew
        rsIDS!ClinicalTrialName = msTrialName
        rsIDS!NextNumberID = NewNextNumberID(msTrialName)
        If mbUsePrefix And msPrefixValue > "" Then
            rsIDS!Prefix = ReplaceQuotes(msPrefixValue)
        Else
            rsIDS!Prefix = Null
        End If
        If mbUseSuffix And msSuffixValue > "" Then
            rsIDS!Suffix = ReplaceQuotes(msSuffixValue)
        Else
            rsIDS!Suffix = Null
        End If
    End If
    ' Store the next number to use
    rsIDS!NextNumber = lNumber + 1
    rsIDS.Update
    rsIDS.Close
    Set rsIDS = Nothing
    
    ' Calculate identifier
    msSubjectIdentifier = msPrefixValue & PadNumber(lNumber, mnNumberWidth) & msSuffixValue
    
    ' Update the RSSubjectIdentifier table
    rsSubjectId!SubjectIdentifier = ReplaceQuotes(msSubjectIdentifier)
    rsSubjectId.Update
    rsSubjectId.Close
    Set rsSubjectId = Nothing
    
    SetReturnMessage eRegResult.RegOK
    
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "GenerateIdentifier", "clsRSSubjectNumbering")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Function

'----------------------------------------------------------------------------------------'
Private Function NewNextNumberID(ByVal sTrialName) As Long
'----------------------------------------------------------------------------------------'
' Generate a unique identifier for a record in the RSNextNumber table
' for this TrialName
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsIDS As ADODB.Recordset

    On Error GoTo ErrHandler
    
    sSQL = "SELECT MAX(NextNumberID) AS MaxID FROM RSNextNumber WHERE " _
            & GetSQLStringEquals("ClinicalTrialName", sTrialName)
    Set rsIDS = New ADODB.Recordset
    rsIDS.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    If IsNull(rsIDS!MaxID) Then
        ' No records yet
        NewNextNumberID = 1
    Else
        NewNextNumberID = rsIDS!MaxID + 1
    End If
    
    rsIDS.Close
    Set rsIDS = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "NewNextNumberID(" & sTrialName & ")", _
                        "clsRSSubjectNumbering")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
            
End Function

'----------------------------------------------------------------------------------------'
Private Function PadNumber(ByVal lNumber As Long, ByVal nWidth As Integer) As String
'----------------------------------------------------------------------------------------'
' Pad the number with leading zeros to the given width
' If nWidth is 0 or less than current number width, leave number unchanged
'----------------------------------------------------------------------------------------'
Dim n As Integer
Dim sFormat As String

    On Error GoTo ErrHandler
    
    If nWidth > 0 Then
        ' Generate format string with requisite no. of 0's
        sFormat = String(nWidth, "0")
        PadNumber = Format(lNumber, sFormat)
    Else      ' Free width
        PadNumber = CStr(lNumber)
    End If

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                        "PadNumber(" & lNumber & "," & nWidth & ")", _
                        "clsRSSubjectNumbering")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'
' Initialise values in the class
'----------------------------------------------------------------------------------------'
    
    ' Do what's necessary to kick MACRO off
    Call InitialisationSettings
    
    msTrialName = ""
    msSite = ""
    mlPersonId = 0
    msPrefixValue = ""
    msSuffixValue = ""

    mlStartNumber = 0
    mnNumberWidth = 0
    mbUsePrefix = False
    mbUseSuffix = False
    msSubjectIdentifier = ""

    msChecksTimeStamp = ""
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_Terminate()
'----------------------------------------------------------------------------------------'

    Set mcolSubjectChecks = Nothing
    Set goUser = Nothing
    
    ' Shut down MACRO database connections
    Call TerminateAllADODBConnections
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub SetReturnMessage(nReturnCode As eRegResult)
'----------------------------------------------------------------------------------------'
' Set the return code and an appropriate return message
'----------------------------------------------------------------------------------------'

    mnResultCode = nReturnCode
    Select Case nReturnCode
    Case eRegResult.RegOK
        msResultMsg = "Registration successful"
    Case eRegResult.RegAlreadyRegistered
        msResultMsg = "Subject has already been registered"
    Case eRegResult.RegMissingInfo
        msResultMsg = "Study, Site and PersonId are not all known"
    Case eRegResult.RegNotUnique
        msResultMsg = "Subject's check values are not unique for this study"
    Case eRegResult.RegError
        msResultMsg = "An error occurred during registration"
    End Select

End Sub
