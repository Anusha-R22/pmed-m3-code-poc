Attribute VB_Name = "modAPIUtils"

'------------------------------------------------------------------
' File: modAPIUtils.bas
' Copyright InferMed Ltd 2004 All Rights Reserved
' Author: Nicky Johns, Feb 2004
' Purpose: Supporting routines for the MACRO API (based on original LISA interface)
'------------------------------------------------------------------
' REVISIONS
' NCJ 14 Aug 2003 - Initial development
' NCJ 27 Aug 03 - Option to load a subject as read-only when retrieving data
' NCJ 21 Jan 04 - Added CreateSubject
' NCJ 2 Feb 04 - This file created from original modLISAUtils.cls
' TA 24/05/2005 - Allow Trialid to be used; allow timestamps
' NCJ 9 Aug 06 - Allow numeric subject labels in GetIDsFromNames
' NCJ 17 Aug 06 - Allow Lab code to be specified for eForm
' NCJ 10 Mar 08 - User and Security DB checking moved here from APILogin
' NCJ 17 Mar 08 - Must set CursorLocation for connections to ensure Oracle works OK
'------------------------------------------------------------------

Option Explicit

' The XML tags
Public Const gsTAG_VISIT As String = "Visit"
Public Const gsTAG_EFORM As String = "Eform"
Public Const gsTAG_QUESTION As String = "Question"

' The XML attributes
Public Const gsATTR_CODE As String = "Code"
Public Const gsATTR_CYCLE As String = "Cycle"
Public Const gsATTR_VALUE As String = "Value"
'TA 28/04/2005 - user may want to set response timestamps
Public Const gsATTR_TIMEZONE As String = "Timezone"
Public Const gsATTR_TIMESTAMP As String = "Timestamp"
' NCJ 17 Aug 06 - Allow setting of eForm lab
Public Const gsATTR_LAB As String = "Lab"

Public Const gsATTR_STUDY As String = "Study"
Public Const gsATTR_SITE As String = "Site"
Public Const gsATTR_LABEL As String = "Label"

' The initial XML for an empty subject
Public Const gsXML_VERSION_HEADER As String = "<?xml version=""1.0""?>"
Public Const gsXML_EMPTY_SUBJ As String = "<MACROSubject> </MACROSubject>"

Public Const gsXMLTAG0_INPUTERR = "<MACROInputErrors>"
Public Const gsXMLTAG1_INPUTERR = "</MACROInputErrors>"



'---------------------------------------------------------------------
Public Function LoadSubject(oUser As MACROUser, _
                        ByVal sStudyName As String, _
                        ByVal sSite As String, _
                        ByVal sSubjLabel As String, _
                        ByRef sErrMsg As String, _
                        Optional bReadOnly As Boolean = False) As StudySubject
'---------------------------------------------------------------------
' Load the subject as specified
' Returns Nothing if subject load not successful
' Load as Read-Only (with no AREZZO) if bReadOnly = TRUE
'---------------------------------------------------------------------
Dim sTempPath As String
Dim oSubject As StudySubject
Dim oArezzo As Arezzo_DM
Dim oStudyDef As StudyDefRO
Dim lStudyId As Long
Dim lSubjId As Long
Dim enUpdateMode As eUIUpdateMode

    On Error GoTo ErrLabel
    
    Set LoadSubject = Nothing
    
    ' Get settings file (TRUE means look one level up because we're in a DLL)
    Call InitialiseSettingsFile(True)
    sTempPath = GetMACROSetting("Temp", App.Path & "\..\Temp\")

    ' Get the study ID and Subject ID (needed for loading)
    If GetIDsFromNames(oUser, sStudyName, sSite, sSubjLabel, lStudyId, lSubjId) Then
        If Not bReadOnly Then
            enUpdateMode = eUIUpdateMode.Read_Write
        Else
            enUpdateMode = eUIUpdateMode.Read_Only
        End If
        Set oArezzo = InitNewArezzo(sTempPath, oUser.CurrentDBConString, lStudyId)
        Set oStudyDef = LoadStudyDef(oUser.CurrentDBConString, lStudyId, oArezzo, sErrMsg)
        If Not oStudyDef Is Nothing Then
            Call oStudyDef.LoadSubject(sSite, lSubjId, oUser.UserName, enUpdateMode, _
                                        oUser.UserNameFull, oUser.UserRole)
            If oStudyDef.Subject.CouldNotLoad Then
                ' Give up
                sErrMsg = "Unable to open subject: " & oStudyDef.Subject.CouldNotLoadReason
            Else
                ' Successfully loaded
                Set LoadSubject = oStudyDef.Subject
            End If
        End If
    Else
        ' Couldn't recognise the subject
        sErrMsg = "Subject does not exist"
    End If
    
    Set oSubject = Nothing
    Set oArezzo = Nothing
    Set oStudyDef = Nothing

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modAPIUtils.LoadSubject"

End Function

'---------------------------------------------------------------------
Public Function InitNewArezzo(ByVal sTempPath As String, _
                            ByVal sDBConString As String, _
                            ByVal lStudyId As Long) As Arezzo_DM
'---------------------------------------------------------------------
' Create and initialise a new instance of AREZZO for the given study
'---------------------------------------------------------------------
Dim oArezzoMemory As clsAREZZOMemory
Dim oArezzo As Arezzo_DM

    On Error GoTo ErrLabel
    
    'Create and initialise a new Arezzo instance
    Set oArezzo = New Arezzo_DM

    Set oArezzoMemory = New clsAREZZOMemory
    Call oArezzoMemory.Load(lStudyId, sDBConString)
    Call oArezzo.Init(sTempPath, oArezzoMemory.AREZZOSwitches)
    Set oArezzoMemory = Nothing

    Set InitNewArezzo = oArezzo
    Set oArezzo = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modAPIUtils.InitNewArezzo"

End Function


'---------------------------------------------------------------------
Private Function LoadStudyDef(ByVal sDBConString As String, _
                            ByVal lStudyId As Long, _
                            oArezzo As Arezzo_DM, _
                            ByRef sMsg As String) As StudyDefRO
'---------------------------------------------------------------------
' Load and return the study def
' Returns Nothing if study not successfully loaded, with sMsg a suitable message
'---------------------------------------------------------------------
Dim sErrMsg As String
Dim bLoadedOK As Boolean
Dim oStudyDef As StudyDefRO

    On Error GoTo ErrLabel
    
    Set oStudyDef = New StudyDefRO
    sErrMsg = oStudyDef.Load(sDBConString, lStudyId, 1, oArezzo)
    If sErrMsg > "" Then
        ' Give up
        sMsg = "Unable to load study: " & sErrMsg
        Set oStudyDef = Nothing
    End If
    
    Set LoadStudyDef = oStudyDef
    Set oStudyDef = Nothing

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modAPIUtils.LoadStudyDef"

End Function

'--------------------------------------------------------------------
Private Function GetIDsFromNames(oUser As MACROUser, _
                        ByVal sStudyName As String, _
                        ByVal sSite As String, _
                        ByVal sSubjLabel As String, _
                        ByRef lStudyId As Long, _
                        ByRef lSubjId As Long) As Boolean
'--------------------------------------------------------------------
' Get the study ID and subject ID of the specified subject
' sSite can be "", with sSubjLabel the subject label
' or if Site > "", sSubjLabel can be either label or numeric ID
' Returns FALSE if no subject found
'--------------------------------------------------------------------
Dim vSubjects As Variant
    
    On Error GoTo ErrLabel
    
    'TA 24/05/2005
    'assume a numeric studyname is studyid
    If IsNumeric(sStudyName) Then
        'convert to the name (inefficient I know - but a quick fix)
        sStudyName = oUser.Studies.StudyById(sStudyName).StudyName
    End If
    
    ' NCJ 9 Aug 06 - Initialise to Null
    vSubjects = Null
    
    If sSite > "" Then
        ' We have a site
        If IsNumeric(sSubjLabel) Then
            ' Treat a numeric subject label as a subject ID
            vSubjects = oUser.DataLists.GetSubjectList(, sStudyName, sSite, Val(sSubjLabel))
        End If
        ' NCJ 9 Aug 06 - Try again if numeric SubjLabel didn't retrieve anything
        ' Did we get a subject this way?
        If IsNull(vSubjects) Then
            vSubjects = oUser.DataLists.GetSubjectList(sSubjLabel, sStudyName, sSite)
        End If
    Else
        ' We don't have a site - assume the label's OK
        vSubjects = oUser.DataLists.GetSubjectList(sSubjLabel, sStudyName)
    End If
    
    If IsNull(vSubjects) Then
        ' Subject does not exist
        lStudyId = 0
        lSubjId = 0
        GetIDsFromNames = False
'        AddErrorMsg eDataInputError.SubjectNotExist, "Subject does not exist"
    Else
        ' Assume subject is in row 0 - Pick off IDs
        lStudyId = vSubjects(eSubjectListCols.StudyId, 0)
        lSubjId = vSubjects(eSubjectListCols.SubjectId, 0)
        GetIDsFromNames = True
    End If

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modAPIUtils.GetIDsFromNames"

End Function

'---------------------------------------------------
Public Function SetXMLRequest(ByVal sXMLData As String, _
                    oXMLDoc As MSXML2.DOMDocument, _
                    ByRef sStudyName As String, _
                    ByRef sSite As String, _
                    ByRef sSubjLabel As String) As Boolean
'---------------------------------------------------
' Initialise oXMLDoc and load in the XML Data string
' Return the Study, Site and Subject details
' Return FALSE if any of the values are missing or if it's invalid XML
'---------------------------------------------------

    On Error GoTo ErrLabel
    
    SetXMLRequest = False
    
    sStudyName = ""
    sSite = ""
    sSubjLabel = ""
    
    Set oXMLDoc = New MSXML2.DOMDocument
    If oXMLDoc.loadXML(sXMLData) Then
        With oXMLDoc.documentElement
            sStudyName = .getAttribute(gsATTR_STUDY)
            sSite = .getAttribute(gsATTR_SITE)
            sSubjLabel = .getAttribute(gsATTR_LABEL)
        End With
        SetXMLRequest = ((sStudyName > "") And (sSite > "") And (sSubjLabel > ""))
    End If

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modAPIUtils.SetXMLRequest"

End Function

'---------------------------------------------------
Public Sub GetCodeAndCycle(oElNode As MSXML2.IXMLDOMElement, _
                            ByRef sCode As String, _
                            ByRef nCycle As Integer)
'---------------------------------------------------
' Retrieve the Code and Cycle attributes from this element
' Code is returned as LOWER CASE
' Code must be present, but Cycle defaults to 1 if not present
'---------------------------------------------------
Dim vCycle As Variant

    On Error GoTo ErrLabel
    
    sCode = LCase(oElNode.getAttribute(gsATTR_CODE))   ' Assume there IS one!
    vCycle = oElNode.getAttribute(gsATTR_CYCLE)
    ' If no cycle assume 1
    If IsNull(vCycle) Then
        nCycle = 1
    Else
        ' Assume a valid integer!
        nCycle = CInt(vCycle)
    End If
                            
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modAPIUtils.GetCodeAndCycle"

End Sub

'---------------------------------------------------------------------
Public Function CreateSubject(sSerialisedUser As String, _
                        ByVal lStudyId As Long, _
                        ByVal sSite As String, _
                        ByRef sErrMsg As String _
                        ) As Long
'---------------------------------------------------------------------
' TA Nov 03
' Create a new subject
' Returns SubjectID if successful, or -1 if not
'---------------------------------------------------------------------
Dim sTempPath As String
Dim oSubject As StudySubject
Dim oArezzo As Arezzo_DM
Dim oStudyDef As StudyDefRO
Dim sCountry As String
Dim oUser As MACROUser

    On Error GoTo ErrLabel
    
    ' Create the MACRO User
    Set oUser = New MACROUser
    Call oUser.SetStateHex(sSerialisedUser)
    
    CreateSubject = -1
    
    ' Get settings file (TRUE means look one level up because we're in a DLL)
    Call InitialiseSettingsFile(True)
    sTempPath = GetMACROSetting("Temp", App.Path & "\..\Temp\")

    Set oArezzo = InitNewArezzo(sTempPath, oUser.CurrentDBConString, lStudyId)
    Set oStudyDef = LoadStudyDef(oUser.CurrentDBConString, lStudyId, oArezzo, sErrMsg)
    If Not oStudyDef Is Nothing Then
    
        sCountry = oUser.GetAllSites.Item(sSite).CountryName
        Call oStudyDef.NewSubject(sSite, oUser.UserName, sCountry, oUser.UserNameFull, oUser.UserRole)
        If oStudyDef.Subject.CouldNotLoad Then
            ' Give up
            sErrMsg = "Unable to create subject: " & oStudyDef.Subject.CouldNotLoadReason
        Else
            ' Successfully loaded
            CreateSubject = oStudyDef.Subject.PersonId
        End If
        
        ' Tidy up objects
        Call Terminate(oStudyDef)
    End If
    
    
    Set oSubject = Nothing
    Set oArezzo = Nothing
    Set oStudyDef = Nothing
    Set oUser = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modAPIUtils.CreateSubject"

End Function

'---------------------------------------------------------------------
Public Sub gLog(sUserName As String, sTaskId As String, sMessage As String, conSec As ADODB.Connection)
'---------------------------------------------------------------------
'REM 11/10/02
'used to write to the LoginLog table
'---------------------------------------------------------------------
' REVISIONS
' DPH 23/03/2004 - objects & connections tidying
'---------------------------------------------------------------------
Dim sSQL As String
Dim sSQLNow As String
Dim nLogNumber As Long
Dim nTimeZone As Integer
Dim sLocation As String
Dim rsLoginLog As ADODB.Recordset
Dim oTimezone As Timezone
Dim SecurityCon As ADODB.Connection

    On Error GoTo ErrHandler
    
    Set oTimezone = New Timezone
    
'    If conSec Is Nothing Then
'        Set SecurityCon = mconSecurity
'    Else
        Set SecurityCon = conSec
'    End If
    
    'Use standard SQL datestamp
    sSQLNow = LocalNumToStandard(IMedNow)
    
    'Log messages have a combined key of LogDateTime and LogNumber. The first log
    'messages for a particular time will have a LogNumber of 0, the next 1 and so on
    'until the LogDateTime moves on a second.
    sSQL = "SELECT LogNumber From LoginLog WHERE LogDateTime = " & sSQLNow

    'assess the number of records and set the LogNumber for this entry (nLogNumber)
    Set rsLoginLog = New ADODB.Recordset
    'Note use of adOpenKeyset cusor. Recordcount does not work with a adOpenDynamic cursor
    rsLoginLog.Open sSQL, SecurityCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    nLogNumber = rsLoginLog.RecordCount
    rsLoginLog.Close
    Set rsLoginLog = Nothing
    
    'truncate large messages that might have ben created by an error
    'to 255 characters so that it fits into field LogMessage in table LoginLog
    If Len(sMessage) > 255 Then
        sMessage = Left(sMessage, 255)
    End If
    
    'check log message for single quotes
    sMessage = ReplaceQuotes(sMessage)
    
    'get time zone off-set
    nTimeZone = oTimezone.TimezoneOffset
    
    'add in location, will always be local
    sLocation = "Local"
    
    'add log
    sSQL = " INSERT INTO LoginLog " _
        & "(LogDateTime,LogNumber,TaskId,LogMessage,UserName,LogDateTime_TZ,Location,Status)" _
        & " Values (" & sSQLNow & "," & nLogNumber & ",'" & sTaskId _
        & "','" & sMessage & "','" & sUserName & "'," & nTimeZone & ",'" & sLocation & "'," & 0 & ")"

    SecurityCon.Execute sSQL
    
    Set oTimezone = Nothing
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "UserLogin.gLog"
End Sub

'---------------------------------------------------
Private Sub Terminate(oStudyDef As StudyDefRO)
'---------------------------------------------------

    Dim oArezzo As Arezzo_DM
    
    If Not oStudyDef Is Nothing Then
        If Not oStudyDef.Subject Is Nothing Then
            'clear up subject
            Call oStudyDef.RemoveSubject
        End If
        Set oArezzo = oStudyDef.Arezzo_DM
        ' Clear up the study def
        Call oStudyDef.Terminate
         ' clear AREZZO (shut down ALM)
        If Not oArezzo Is Nothing Then
            oArezzo.Finish
            Set oArezzo = Nothing
        End If
    End If
End Sub

'---------------------------------------------------
Public Function CheckLab(sDBCon As String, sLabCode As String, sSite As String) As Boolean
'---------------------------------------------------
' NCJ 18 Aug 06 - Check that a lab is valid for this site
' sLabCode is lower case lab code
' Code copied from clsLabs/clsLab (which we can't use here because they reference too many other things!)
'---------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bLabOK As Boolean

    On Error GoTo ErrLabel
    
    bLabOK = False
    ' Get labs for this site
    sSQL = "SELECT Laboratory.LaboratoryCode " _
                & " FROM Laboratory, SiteLaboratory " _
                & " WHERE Laboratory.LaboratoryCode = SiteLaboratory.LaboratoryCode " _
                & " AND SiteLaboratory.Site = '" & sSite & "'" _
                & " ORDER BY Laboratory.LaboratoryCode"
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, sDBCon
    
    ' Check that the given lab is in the list
    Do While Not rsTemp.EOF And Not bLabOK
        If LCase(rsTemp.Fields!LaboratoryCode) = sLabCode Then
            bLabOK = True
        End If
        rsTemp.MoveNext
    Loop
    
    rsTemp.Close
    Set rsTemp = Nothing

    CheckLab = bLabOK

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modAPIUtils.CheckLab"
End Function

'--------------------------------------------------------------------------------------------------
Public Function GetSecurityCon(Optional sSecCon As String = "") As String
'--------------------------------------------------------------------------------------------------
' Function returns the connection string for the MACRO security DB
' NCJ 23 Nov 07 - If sSecCon is non-empty, use this as connection string
' otherwise use the one specified in the settings file
' Returns "" if invalid Security DB
'--------------------------------------------------------------------------------------------------
Dim sSec As String
Dim oconSecurity As Connection
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    ' Trap errors and return ""
    On Error GoTo SecError
    
    sSec = ""
    If sSecCon <> "" Then
        ' See if this connection string looks OK
        sSec = sSecCon
        Set oconSecurity = New Connection
        oconSecurity.Open sSec
        ' Check that it looks like a MACRO Security DB
        sSQL = "SELECT * FROM SECURITYCONTROL"
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open sSQL, oconSecurity, adOpenKeyset, adLockReadOnly, adCmdText
        rsTemp.Close
        Set rsTemp = Nothing
        oconSecurity.Close
        Set oconSecurity = Nothing
    Else
        ' Get the default database
        InitialiseSettingsFile True
        sSec = GetMACROSetting(MACRO_SETTING_SECPATH, "")
        If sSec <> "" Then
            ' Assume it's a valid connection
            sSec = DecryptString(sSec)
        End If
    End If
    ' If we get here, we're OK!
    GetSecurityCon = sSec
    
Exit Function

SecError:
    ' Error in getting or verifying Security DB, so return ""
    GetSecurityCon = ""
End Function

'----------------------------------------------------------------------------------------'
Public Function UserExists(sSecCon As String, ByRef sUserName As String) As Integer
'----------------------------------------------------------------------------------------'
' NCJ 25 Feb 08 - Check the existence of a UserName
' Returns 0 if user name exists AND is enabled
' Returns 1 if user name exists but account is disabled
' Returns 2 if user name does not exist
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsUser As ADODB.Recordset
Dim oSecCon As Connection

    On Error GoTo Errorlabel
    
    ' Default to non-existent
    UserExists = 2
    
    ' Assume security connection is valid
    Set oSecCon = New Connection
    oSecCon.Open sSecCon
    ' Must set CursorLocation otherwise Oracle won't work properly
    oSecCon.CursorLocation = adUseClient
        
    ' Get UserName from the MacroUser table
    ' Check user name in Uppercase in Oracle as it is case sensitive
    Select Case Connection_Property(CONNECTION_PROVIDER, sSecCon)
    Case CONNECTION_MSDAORA, CONNECTION_ORAOLEDB_ORACLE
        sSQL = "SELECT * FROM MACROUser WHERE NLS_UPPER(UserName) = NLS_UPPER('" & sUserName & "')"
    Case Else
        sSQL = "SELECT * FROM MACROUser WHERE UserName = '" & sUserName & "'"
    End Select
    
    Set rsUser = New ADODB.Recordset
    rsUser.Open sSQL, oSecCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsUser.RecordCount > 0 Then
        ' User does exist
        If rsUser!Enabled = 0 Then
            ' Disabled
            UserExists = 1
        Else
            ' All OK
            UserExists = 0
        End If
    Else
        ' User does not exist
        UserExists = 2
    End If
    
    rsUser.Close
    Set rsUser = Nothing
    
    oSecCon.Close
    Set oSecCon = Nothing
    
Exit Function
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APILogin.UserExists"
End Function

'----------------------------------------------------------------------------------------'
Public Function DBExists(sSecCon As String, sDBCode As String) As Boolean
'----------------------------------------------------------------------------------------'
' NCJ 12 Mar 08 - Check that the given database is registered in this SecurityDB
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim oSecCon As Connection

    On Error GoTo Errorlabel

    ' Assume security connection is valid
    Set oSecCon = New Connection
    oSecCon.Open sSecCon
        
    ' Check DBCode in Uppercase in Oracle as it is case sensitive
    Select Case Connection_Property(CONNECTION_PROVIDER, sSecCon)
    Case CONNECTION_MSDAORA, CONNECTION_ORAOLEDB_ORACLE
        sSQL = "SELECT * FROM DATABASES WHERE upper(DATABASECODE) = upper('" & sDBCode & "')"
    Case Else
        sSQL = "SELECT * FROM DATABASES WHERE DATABASECODE = '" & sDBCode & "'"
    End Select
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, oSecCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    ' Did we find it?
    DBExists = (rsTemp.RecordCount > 0)
    
    rsTemp.Close
    oSecCon.Close
    Set rsTemp = Nothing
    Set oSecCon = Nothing
    
Exit Function
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "APILogin.DBExists"
End Function
