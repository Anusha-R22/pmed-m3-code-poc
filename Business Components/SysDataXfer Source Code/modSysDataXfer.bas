Attribute VB_Name = "modSysDataXfer"
Option Explicit
'--------------------------------------------------------------------------------------------------
' REVISIONS
'   DPH 23/01/2003 Added Getting / Saving report files
'   REM 13/10/03 - Pass last transfer date into routine CreateReportFilesZIP (sent from site), no longer get it from server Logdetails table
'   REM 03/12/03 - In routine WriteMsgToMessageTable Added ReplaceQuotes around sMessageParameters in the Insert statment
'   REM 16/02/04 - When AddSystemMessage routine is used there were some cases where the optional parameters were incorrectly set
'   REM 17/03/04 - In routine GetLogDetailsMessages added a replace function to replace any "|" or "*" characters with the words "PIPE" or "STAR"
'   REM 23/03/04 - Added LocalNumToStandard to dates in functions GetUserLogMessages, SetLoginLogStatus
'   rem/ic 21/05/2004 remove '' from around nHistoryNumber and sPasswordCreated in WriteChangePassword()
'   MLM 21/06/05: bug 2570: Added commas to SQL
'                 bug 2574: Added ReplaceQuotes() around UserNameFulls
'   TA  18/01/2006 - MessageId now calculated by a sequence to avoid duplicate id problem
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
Public Function CreateSecurityConnection(ByRef sErrorMsg As String, Optional sSecurityCon As String = "") As ADODB.Connection
'--------------------------------------------------------------------------------------------------
'REM 18/11/02
'Create the security connection, will used passed in Security connection string but if none will use registry value
'--------------------------------------------------------------------------------------------------
Dim sSecCon As String
Dim oSecCon As ADODB.Connection

    On Error GoTo ErrLabel

    InitialiseSettingsFile (True)

    If sSecurityCon = "" Then
        sSecCon = SecurityDatabasePath
    Else
        sSecCon = sSecurityCon
    End If

    Set oSecCon = New ADODB.Connection
    oSecCon.Open sSecCon
    oSecCon.CursorLocation = adUseClient
    
    sSecurityCon = sSecCon
    
    Set CreateSecurityConnection = oSecCon
   
Exit Function
ErrLabel:
    sErrorMsg = "Security Database connection error: " & Err.Description & ": Error no. " & Err.Number
End Function

'----------------------------------------------------------------------------------------'
Public Property Get SecurityDatabasePath() As String
'----------------------------------------------------------------------------------------'

    SecurityDatabasePath = GetMACROSetting("SecurityPath", DefaultSecurityDatabasePath)
    If SecurityDatabasePath <> "" Then
        SecurityDatabasePath = DecryptString(SecurityDatabasePath)
    End If

End Property

'----------------------------------------------------------------------------------------'
Private Function DefaultSecurityDatabasePath() As String
'----------------------------------------------------------------------------------------'
' Get MACRO's default security path (i.e. the one set on installation)
'----------------------------------------------------------------------------------------'

    DefaultSecurityDatabasePath = ""

End Function

'----------------------------------------------------------------------------------------'
Public Function WriteUserDetails(conMACRO As ADODB.Connection, oSecCon As ADODB.Connection, sSecCon As String, _
                                 sTrialSite As String, nMessageType As Integer, sSystemUserName As String, _
                                 ByVal sMessageParameters As String, sMessageBody As String, _
                                 ByRef nMessageReceived As Integer, ByRef sErrMessage As String) As Boolean
'----------------------------------------------------------------------------------------'
'REM 11/11/02
'Writes user details to the Secuirty Database MACROUser table
' MLM 21/06/05: bug 2570: Added commas to SQL
'               bug 2574: Added ReplaceQuotes() around UserNameFulls
'----------------------------------------------------------------------------------------'
Dim sSql As String
Dim rsUser As ADODB.Recordset
Dim vMessage As Variant
Dim vUserDetails As Variant
Dim sUserName As String
Dim sUserNameFull As String
Dim sUserPassword As String
Dim nEnabled As Integer
Dim sLastLogin As String
Dim sFirstLogin As String
Dim sDefaultUserRole As String
Dim nFailedAttempts As Integer
Dim sPswdCreated As String
Dim sSiteCode As String
Dim bPendingMessage As Boolean
Dim nNewUser As Integer
Dim nSysAdmin As Integer
Dim oUser As MACROUser
Dim sSite As String
Dim bLockoutUser As Boolean
Dim vUserSites As Variant
Dim sUserSite As String
Dim bUserSite As Boolean
Dim i As Integer
Dim rsConfUser As ADODB.Recordset
Dim nUser As Integer

    On Error GoTo ErrLabel
    
    vUserDetails = Split(sMessageParameters, gsPARAMSEPARATOR)
    
    sUserName = vUserDetails(0)
    sUserNameFull = vUserDetails(1)
    sUserPassword = vUserDetails(2)
    nEnabled = Val(vUserDetails(3))
    sFirstLogin = vUserDetails(4)
    sLastLogin = vUserDetails(5)
    nFailedAttempts = Val(vUserDetails(6))
    sDefaultUserRole = "MACROUser"
    sPswdCreated = vUserDetails(7)
    nSysAdmin = vUserDetails(8)
    nNewUser = vUserDetails(9)
    
    'REM 21/01/03 - check user name in Uppercase in Oracle as it is case sensitive
    Select Case Connection_Property(CONNECTION_PROVIDER, sSecCon)
    Case CONNECTION_MSDAORA, CONNECTION_ORAOLEDB_ORACLE
        sSql = "SELECT COUNT (*) FROM MACROUser WHERE upper(UserName) = '" & UCase(sUserName) & "'"
    Case Else
        sSql = "SELECT COUNT (*) FROM MACROUser WHERE UserName = '" & sUserName & "'"
    End Select
'    sSQL = "SELECT COUNT (*) FROM MACROUser WHERE UserName = '" & sUserName & "'"
    Set rsUser = New ADODB.Recordset
    rsUser.Open sSql, oSecCon, adOpenKeyset, adLockReadOnly, adCmdText
    nUser = rsUser.Fields(0)
        
    rsUser.Close
    Set rsUser = Nothing
    
    'get the trialsite name, will be "server" if finds nothing in MACRODBsetting table
    sSiteCode = GetDBSettings(conMACRO, "datatransfer", "dbsitename", gsSERVER)
    
    bLockoutUser = False
    
    'if user name exists then do an update
    If nUser <> 0 Then
        Select Case nNewUser
        Case eUserDetails.udNewUser 'a new user but one with the same name already exists, so lock user out and send message to all sites the user is on
            sSql = "UPDATE MACROUser SET Enabled = " & eUserStatus.usDisabled _
                & " WHERE UserName = '" & sUserName & "'"
            oSecCon.Execute sSql
            
            'get all user details of user that the site user name conflicted with as this will be used to over
            'write the site user so the same user name does not have two different user details
            sSql = "SELECT * FROM MACROUser WHERE UserName = '" & sUserName & "'"
            Set rsConfUser = New ADODB.Recordset
            rsConfUser.Open sSql, oSecCon, adOpenKeyset, adLockReadOnly, adCmdText
            
            bLockoutUser = True
            
            'set new message parameters i.e. Disabled user
            sMessageParameters = sUserName & gsPARAMSEPARATOR & rsConfUser!UserNameFull & gsPARAMSEPARATOR & rsConfUser!UserPassword & gsPARAMSEPARATOR & eUserStatus.usDisabled & gsPARAMSEPARATOR & rsConfUser!LastLogin & gsPARAMSEPARATOR & rsConfUser!FirstLogin & gsPARAMSEPARATOR & rsConfUser!FailedAttempts & gsPARAMSEPARATOR & rsConfUser!PasswordCreated & gsPARAMSEPARATOR & rsConfUser!SysAdmin & gsPARAMSEPARATOR & eUserDetails.udDisableUser
            
            rsConfUser.Close
            Set rsConfUser = Nothing
            
            'set to false so it will write new message
            bPendingMessage = False
            
            'log the disabling of the user on the server log
            Set oUser = New MACROUser
            Call oUser.gLog(sSystemUserName, "UserNameConflict", "User " & sUserName & " was disabled due to a conflict in User Name with a user on Site " & sTrialSite, oSecCon)
            Set oUser = Nothing
            
        Case eUserDetails.udEditUser
    
            'Check if there are any messages pending for this user, if so do not update user
            bPendingMessage = CheckPendingMessages(conMACRO, sTrialSite, sUserName, nMessageType, sSiteCode)
            
            If Not bPendingMessage Then
                sSql = "UPDATE MACROUser SET UserNameFull = '" & ReplaceQuotes(sUserNameFull) & "', " _
                    & " Enabled = " & nEnabled & ", " _
                    & " FailedAttempts = " & nFailedAttempts & "," _
                    & " SysAdmin = " & nSysAdmin _
                    & " WHERE UserName = '" & sUserName & "'"
                oSecCon.Execute sSql
            End If
            
        Case eUserDetails.udDisableUser 'disable due to user name conflict
        
            ' MLM 21/06/05: bug 2570: Corrected commas in SQL
            sSql = "UPDATE MACROUser SET UserNameFull = '" & ReplaceQuotes(sUserNameFull) & "', " _
                    & " UserPassword = '" & sUserPassword & "'," _
                    & " Enabled = " & eUserStatus.usDisabled & "," _
                    & " LastLogin = " & sLastLogin & ", " _
                    & " FirstLogin = " & sFirstLogin & ", " _
                    & " FailedAttempts = " & nFailedAttempts & "," _
                    & " PasswordCreated = " & sPswdCreated & ", " _
                    & " SysAdmin = " & nSysAdmin _
                    & " WHERE UserName = '" & sUserName & "'"
            oSecCon.Execute sSql
            
            'log the disabling of the user on the site log
            Set oUser = New MACROUser
            Call oUser.gLog(sSystemUserName, "UserNameConflict", "User " & sUserName & " was disabled due to a conflict in User Name on the server", oSecCon)
            Set oUser = Nothing
        
        End Select

    Else 'else create a new user, no need therefore to check for pending message
    
        sSql = "INSERT INTO MACROUser VALUES ('" & sUserName & "','" & ReplaceQuotes(sUserNameFull) & "','" & sUserPassword & "'," _
                                        & nEnabled & "," & sFirstLogin & "," & sLastLogin & ",'" & sDefaultUserRole _
                                        & "'," & nFailedAttempts & "," & sPswdCreated & "," & nSysAdmin & ")"
        oSecCon.Execute sSql

        'no pending messages if its a new user
        bPendingMessage = False
    End If
    
    'don't add system message if there is a pending message already
    If Not bPendingMessage Then
        'only distrabute user details if its a server db
        If LCase(sSiteCode) = gsSERVER Then
            
            'if user has been locked out due to user name conflict then
            If bLockoutUser Then
            
'                'check if user has been assigned to any sites, if not just send message back to site it came from (i.e. sTrialSite)
'                vUserSites = UserSites(conMACRO, sUserName)
'                If IsNull(vUserSites) Then
'                    sSite = sTrialSite
'                Else
                    sSite = sTrialSite
                    'add system messages to Server Message table and distribute it back to the site it came from (need to do this in the case of a lockout due to conflicting user names)
                    Call AddSystemMessage(conMACRO, nMessageType, sSystemUserName, sUserName, sMessageBody, sMessageParameters, sSite)
                    'set sSite back to "" so the AddSystemMessage will distrabute the message to all other sites the user is on
                    sSite = ""

'                End If
'            Else
'                sSite = ""
            End If
            
            'add system messages to Server Message table as need to distribute the new User Details to other sites (pass in sTrialSite so it doesn't send the same message back to the site it came from, that is taken care of above if nessesary)
            Call AddSystemMessage(conMACRO, nMessageType, sSystemUserName, sUserName, sMessageBody, sMessageParameters, sSite, "", sTrialSite)
            
        End If
    End If
    
    If bPendingMessage Then
        nMessageReceived = MessageReceived.PendingOverRule 'there is a Pending Message, so mesasge not written away
    Else
        nMessageReceived = MessageReceived.Received
    End If
    
    WriteUserDetails = True

Exit Function
ErrLabel:
    WriteUserDetails = False
    sErrMessage = "Error occurred while writing user details: " & "Error Description: " _
                 & Err.Description & ", Error Number: " & Err.Number & ", SQL = " & sSql
End Function

'----------------------------------------------------------------------------------------'
Public Function WriteUserRole(conMACRO As ADODB.Connection, oSecCon As ADODB.Connection, sTrialSite As String, sDatabaseCode As String, _
                              nMessageType As Integer, sSystemUserName As String, ByVal sMessageParameters As String, _
                              sMessageBody As String, ByRef nMessageReceived As Integer, ByRef sErrorMsg As String) As Boolean
'----------------------------------------------------------------------------------------'
'REM 12/11/02
'Writes UserRole to the MACRO DB UserRole table
'----------------------------------------------------------------------------------------'
Dim sSql As String
Dim rsUserRole As ADODB.Recordset
Dim sUserName As String
Dim vUserRole As Variant
Dim sRoleCode As String
Dim sStudyCode As String
Dim sDBSitecode As String
Dim sSiteCode As String
Dim nTypeOfInstallation As Integer
Dim nAdd As Integer
Dim bAdd As Boolean
Dim sMsgBody As String
Dim sMsgParameters As String
Dim bPendingMessage As Boolean
Dim rsUser As ADODB.Recordset

    On Error GoTo ErrLabel
    
    vUserRole = Split(sMessageParameters, gsPARAMSEPARATOR)
    
    sUserName = vUserRole(0)
    sRoleCode = vUserRole(1)
    sStudyCode = vUserRole(2)
    sSiteCode = vUserRole(3)
    nTypeOfInstallation = vUserRole(4)
    nAdd = vUserRole(5)
    bAdd = (nAdd = 1)
    
    'get the sitecode of the database, will be "server" if finds nothing in MACRODBsetting table or is set as a server
    sDBSitecode = GetDBSettings(conMACRO, "datatransfer", "dbsitename", gsSERVER)
    
    bPendingMessage = CheckPendingMessages(conMACRO, sTrialSite, sUserName, nMessageType, sDBSitecode)
    
    'if there are no pending messages
    If Not bPendingMessage Then
    
        'if it is a UserRole to be added then
        If bAdd Then
            
            'check that the new UserRole to insert does not have primary key conflicts
            sSql = "SELECT COUNT (*) FROM UserRole" _
                & " WHERE UserName = '" & sUserName & "'" _
                & " AND RoleCode = '" & sRoleCode & "'" _
                & " AND StudyCode = '" & sStudyCode & "'" _
                & " AND SiteCode = '" & sSiteCode & "'" _
                & " AND TypeOfInstallation = " & nTypeOfInstallation
            Set rsUserRole = New ADODB.Recordset
            rsUserRole.Open sSql, conMACRO, adOpenKeyset, adLockReadOnly, adCmdText
            
            'If no conflict go ahead, else ignore as already been written to database
            If rsUserRole.Fields(0) = 0 Then
                sSql = "INSERT INTO UserRole " _
                & " VALUES ('" & sUserName & "','" & sRoleCode & "','" & sStudyCode & "','" & sSiteCode & "', " & nTypeOfInstallation & ")"
                conMACRO.Execute sSql, , adCmdText
                
                'Checks to see if there is a User/Database combination in the UserDatabase table in the Security database
                'if not adds it
                Call AddUserDatabase(oSecCon, sUserName, sDatabaseCode)
                
                If LCase(sDBSitecode) = gsSERVER Then
                    'add system messages to Server Message table if need to distribute the new UserRole to other sites
                    Call AddSystemMessage(conMACRO, nMessageType, sSystemUserName, sUserName, sMessageBody, sMessageParameters, sSiteCode, "", sTrialSite)
                End If
                
            End If
            
        Else ' else its a UserRole to be removed
            
            If LCase(sSiteCode) = gsSERVER Then
                'add system messages to Server Message table if need to distribute the Delete of the UserRole to other sites
                Call AddSystemMessage(conMACRO, nMessageType, sSystemUserName, sUserName, sMessageBody, sMessageParameters, sSiteCode, "", sTrialSite)
            End If
            
            sSql = "DELETE FROM UserRole WHERE UserName = '" & sUserName & "'" _
                & " AND RoleCode = '" & sRoleCode & "'" _
                & " AND StudyCode = '" & sStudyCode & "'" _
                & " AND SiteCode = '" & sSiteCode & "'" _
                & " AND TypeOfInstallation = " & nTypeOfInstallation
            conMACRO.Execute sSql, , adCmdText
            
            'Checks to see if there are any UserRoles for the User, if not removes the
            'User/Database from the UserDatabse table in the security database
            Call DeleteUserDatabase(conMACRO, oSecCon, sUserName, sDatabaseCode)
            
        End If
 
    Else 'there are pending UserRole messages to be sent to the site
    
        'swap the adds to delete and the deletes to add so restores UserRole at the site in came from
        If nAdd = eUserRole.urAdd Then
            nAdd = eUserRole.urDelete
            sMsgBody = "Delete User Role"
        Else
            nAdd = eUserRole.urAdd
            sMsgBody = "New User Role"
        End If
        
        'new message parameters
        sMsgParameters = sUserName & gsPARAMSEPARATOR & sRoleCode & gsPARAMSEPARATOR & sStudyCode & gsPARAMSEPARATOR & sSiteCode & gsPARAMSEPARATOR & 1 & gsPARAMSEPARATOR & nAdd
        'add system messages to Server Message table to restore UserRoles on the site they came from as there are pending UserRole messages for the server so
        Call AddSystemMessage(conMACRO, ExchangeMessageType.RestoreUserRole, sSystemUserName, sUserName, sMsgBody, sMsgParameters, sTrialSite)

    End If
    
    'set the Message received depending on whether there was a pending message
    If bPendingMessage Then
        nMessageReceived = MessageReceived.PendingOverRule
    Else
        nMessageReceived = MessageReceived.Received
    End If
    
    WriteUserRole = True
    
Exit Function
ErrLabel:
    WriteUserRole = False
    sErrorMsg = "Error occurred while writing UserRoles: " & "Error Description: " & Err.Description _
                & "Error Number: " & Err.Number & ", SQL = " & sSql
End Function

'----------------------------------------------------------------------------------------'
Public Function WriteChangePassword(conMACRO As ADODB.Connection, oSecCon As ADODB.Connection, sTrialSite As String, _
                                    nMessageType As Integer, sSystemUserName As String, sMessageParameters As String, _
                                    sMessageBody As String, ByRef nMessageRecieved As Integer, ByRef sErrMessage As String) As Boolean
'----------------------------------------------------------------------------------------'
'REM 12/11/02
'Writes the new passoword to the secuirty database
'always set failedattempts to 0 when resetting a users password
'----------------------------------------------------------------------------------------'
Dim sSql As String
Dim rsPassword As ADODB.Recordset
Dim vChangePassword As Variant
Dim sUserName As String
Dim sUserPassword As String
Dim sPasswordCreated As String
Dim sFirstLogin As String
Dim sLastLogin As String
Dim sSiteCode As String
Dim bPendingMessage As Boolean
Dim rsMaxHistoryNo As ADODB.Recordset
Dim nHistoryNumber As Integer

    On Error GoTo ErrLabel

    vChangePassword = Split(sMessageParameters, gsPARAMSEPARATOR)
    
    sUserName = vChangePassword(0)
    sUserPassword = vChangePassword(1)
    'REM 09/12/03 - Swapped lastlogin and first login as they were the wrong way round
    sLastLogin = vChangePassword(2)
    sFirstLogin = vChangePassword(3)
    sPasswordCreated = vChangePassword(4)
    
    'get the trialsite name, will be "server" if finds nothing in MACRODBsetting table
    sSiteCode = GetDBSettings(conMACRO, "datatransfer", "dbsitename", gsSERVER)
    
    bPendingMessage = CheckPendingMessages(conMACRO, sTrialSite, sUserName, nMessageType, sSiteCode)
    
    'check to see if there are any pending messages for this user, site and message type (will return false if its a site database)
    If Not bPendingMessage Then
        sSql = "UPDATE MACROUser SET UserPassword = '" & sUserPassword & "', " _
            & " PasswordCreated = " & sPasswordCreated & "," _
            & " FirstLogin = " & sFirstLogin & "," _
            & " Lastlogin = " & sLastLogin & "," _
            & " FailedAttempts = 0" _
            & " WHERE UserName = '" & sUserName & "'"
        oSecCon.Execute sSql, , adCmdText
        
        'only add system message to distribute user password if its a server database
        If LCase(sSiteCode) = gsSERVER Then
            'add system messages to Server Message table if need to distribute the new User Password details to other sites
            Call AddSystemMessage(conMACRO, nMessageType, sSystemUserName, sUserName, sMessageBody, sMessageParameters, "", "", sTrialSite)
        End If
        
    End If
    
    'Even if there is a pending message for Change Password must still add the New Password to the Password History table so it stays up to date with the site's security database
    'get max passord history number for user
    sSql = "SELECT MAX (HistoryNumber) as MaxHistoryNumber FROM PasswordHistory" _
        & " WHERE UserName = '" & sUserName & "'"
    Set rsMaxHistoryNo = New ADODB.Recordset
    rsMaxHistoryNo.Open sSql, oSecCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    If IsNull(rsMaxHistoryNo!MaxHistoryNumber) Then
        nHistoryNumber = 1
    Else
        nHistoryNumber = rsMaxHistoryNo!MaxHistoryNumber + 1
    End If
    
    'rem/ic 21/05/2004 remove '' from around nHistoryNumber and sPasswordCreated
    'insert new password into PasswordHistory table (don't have to worry about primary key violations as the history number is always unique)
    sSql = "INSERT INTO PasswordHistory VALUES ('" & sUserName & "', " & nHistoryNumber & ", " & sPasswordCreated & ", '" & sUserPassword & "')"
    oSecCon.Execute sSql
    
    'set the Mesasge Received status depending in the Pending message status
    If bPendingMessage Then
        nMessageRecieved = MessageReceived.PendingOverRule
    Else
        nMessageRecieved = MessageReceived.Received
    End If
    
    WriteChangePassword = True

Exit Function
ErrLabel:
    WriteChangePassword = False
    sErrMessage = "Error occurred while writing the changed password: " & "Error Description: " _
                    & Err.Description & "Error Number: " & Err.Number & ", SQL = " & sSql
End Function

'----------------------------------------------------------------------------------------'
Public Function WriteRole(oSecCon As ADODB.Connection, sMessageParameters As String, ByRef sErrMessage As String) As Boolean
'----------------------------------------------------------------------------------------'
'REM 12/11/02
'Writes the role to the Security database
'----------------------------------------------------------------------------------------'
Dim sSql As String
Dim rsRole As ADODB.Recordset
Dim vRole As Variant
Dim sRoleCode As String
Dim sRoleDescription As String
Dim nRoleEnabled As Integer
Dim sFunctionCodes As String
Dim sFunctionCode As String
Dim vFunctionCodes As Variant
Dim nSysAdmin As Integer
Dim i As Integer

    On Error GoTo ErrLabel

    vRole = Split(sMessageParameters, gsPARAMSEPARATOR)
    
    sRoleCode = vRole(0)
    sRoleDescription = vRole(1)
    nRoleEnabled = vRole(2)
    nSysAdmin = vRole(3)
    sFunctionCodes = vRole(4)
        
    'get all the function codes for the role
    vFunctionCodes = Split(sFunctionCodes, gsSEPARATOR)

    'check to see if the rode already exists
    sSql = "SELECT * FROM Role WHERE RoleCode = '" & sRoleCode & "'"
    Set rsRole = New ADODB.Recordset
    rsRole.Open sSql, oSecCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    'if it doesn't then add it to the role table
    If rsRole.RecordCount = 0 Then
        sSql = "INSERT INTO Role VALUES ('" & sRoleCode & "','" & sRoleDescription & "'," & nRoleEnabled & "," & nSysAdmin & ")"
        oSecCon.Execute sSql, , adCmdText
    Else 'else update role
        sSql = "UPDATE Role SET RoleDescription = '" & sRoleDescription & "'," _
            & " SysAdmin = " & nSysAdmin _
            & " WHERE RoleCode = '" & sRoleCode & "'"
        oSecCon.Execute sSql, , adCmdText
    End If
    
    'then delete all from the role function table to do with this role
    sSql = "DELETE FROM RoleFunction WHERE RoleCode ='" & sRoleCode & "'"
    oSecCon.Execute sSql, , adCmdText

    'then loop through all the function codes and add them to the Role function table
    For i = 0 To UBound(vFunctionCodes)
        sFunctionCode = vFunctionCodes(i)
    
        sSql = "INSERT INTO RoleFunction " _
        & " VALUES ('" & sRoleCode & "','" & sFunctionCode & "'" & ")"
        oSecCon.Execute sSql, , adCmdText
    Next

    rsRole.Close
    Set rsRole = Nothing
    
    WriteRole = True

Exit Function
ErrLabel:
    WriteRole = False
    sErrMessage = "Error occurred while writing the Role " & sRoleCode & ": " & "Error Description: " _
                & Err.Description & "Error Number: " & Err.Number & ", SQL = " & sSql
End Function

'----------------------------------------------------------------------------------------'
Public Function WriteSystemLog(conMACRO As ADODB.Connection, sMessageParameters As String, sTrialSite As String, sErrorMsg As String) As Boolean
'----------------------------------------------------------------------------------------'
'REM 12/11/02
'Writes the system log message to the MACRO DB Message table and to the System Log table
'----------------------------------------------------------------------------------------'
Dim sSql As String
Dim sUserName As String
Dim vSysLogs As Variant
Dim vLog As Variant
Dim sLog As String
Dim sLogDateTime As String
Dim nLogNumber As Integer
Dim sTaskId As String
Dim sLogMessage As String
Dim nLogDataTime_TZ As Integer
Dim sLocation As String
Dim nStatus As Integer
Dim rsLog As ADODB.Recordset
Dim i As Integer

    On Error GoTo ErrLabel

    vSysLogs = Split(sMessageParameters, gsSEPARATOR)
    
    'loop through all system log messages
    For i = 0 To UBound(vSysLogs)
    
        sLog = vSysLogs(i)
        
        vLog = Split(sLog, gsPARAMSEPARATOR)
        
        sLogDateTime = vLog(0)
        nLogNumber = vLog(1)
        sTaskId = vLog(2)
        sLogMessage = vLog(3)
        sUserName = vLog(4)
        nLogDataTime_TZ = vLog(5)
        sLocation = vLog(6)
        nStatus = 1
    
        sSql = "SELECT COUNT (*) FROM LogDetails" _
            & " WHERE LOGDATETIME = " & sLogDateTime _
            & " AND LOGNUMBER = " & nLogNumber _
            & " AND TASKID = '" & sTaskId & "'" _
            & " AND LOCATION = '" & sTrialSite & "'"
        Set rsLog = New ADODB.Recordset
        rsLog.Open sSql, conMACRO, adOpenKeyset, adLockReadOnly, adCmdText
        
        'check to see if there is a primary key violation, if so the record has already been entered
        If rsLog.Fields(0) = 0 Then
    
            sSql = "INSERT INTO LogDetails " _
                & " VALUES (" & sLogDateTime & ", " & nLogNumber & ", '" & sTaskId & "', '" & ReplaceQuotes(sLogMessage) & "', '" & sUserName _
                & "', " & nLogDataTime_TZ & ", '" & sTrialSite & "'," & nStatus & ")"
            conMACRO.Execute sSql, , adCmdText
            
        End If
        
    Next
    
    WriteSystemLog = True

Exit Function
ErrLabel:
    WriteSystemLog = False
    sErrorMsg = "Error while writing to User Log. " & "Error Description: " & Err.Description & "Error Number: " _
                & Err.Number & ", SQL = " & sSql
End Function

'----------------------------------------------------------------------------------------'
Public Function WriteUserLog(oSecCon As ADODB.Connection, sMessageParameters As String, sTrialSite As String, ByRef sErrorMsg As String) As Boolean
'----------------------------------------------------------------------------------------'
'REM 12/11/02
'Writes the user log to the Login Log table in the security database
'----------------------------------------------------------------------------------------'
Dim sSql As String
Dim sUserName As String
Dim sLogDateTime As String
Dim nLogNumber As Integer
Dim sTaskId As String
Dim sLogMessage As String
Dim nLogDataTime_TZ As Integer
Dim sLocation As String
Dim vUserLogs As Variant
Dim vLog As Variant
Dim sLog As String
Dim i As Integer
Dim nStatus As Integer
Dim rsLog As ADODB.Recordset

    On Error GoTo ErrLabel
    
    'split up into each User log message
    vUserLogs = Split(sMessageParameters, gsSEPARATOR)
    
    'loop through each message and add it to the LoginLog table
    For i = 0 To UBound(vUserLogs)
        
        sLog = vUserLogs(i)
        
        'split message into indivitual parameters
        vLog = Split(sLog, gsPARAMSEPARATOR)
    
        sLogDateTime = vLog(0)
        nLogNumber = vLog(1)
        sTaskId = vLog(2)
        sLogMessage = vLog(3)
        sUserName = vLog(4)
        nLogDataTime_TZ = vLog(5)
        sLocation = vLog(6)
        nStatus = 1
    
        sSql = "SELECT COUNT (*) FROM LoginLog" _
            & " WHERE LOGDATETIME = " & sLogDateTime _
            & " AND LOGNUMBER = " & nLogNumber _
            & " AND TASKID = '" & sTaskId & "'" _
            & " AND LOCATION = '" & sTrialSite & "'"
        Set rsLog = New ADODB.Recordset
        rsLog.Open sSql, oSecCon, adOpenKeyset, adLockReadOnly, adCmdText
        
        'check to see if there is a primary key violation, if so the record has already been enetered
        If rsLog.Fields(0) = 0 Then
        
            'insert into Login Log
            ' MLM 23/06/05: Use ReplaceQuotes here too, in case the message contains a user name with 's
            sSql = "INSERT INTO LoginLog " _
                & " VALUES (" & sLogDateTime & ", " & nLogNumber & ", '" & sTaskId & "', '" & ReplaceQuotes(sLogMessage) & "', '" & sUserName _
                & "', " & nLogDataTime_TZ & ", '" & sTrialSite & "'," & nStatus & ")"
            oSecCon.Execute sSql, , adCmdText
            
        End If
    Next
    
    WriteUserLog = True

Exit Function
ErrLabel:
    WriteUserLog = False
    sErrorMsg = "Error while writing to User Log. " & "Error Description: " & Err.Description & "Error Number: " _
                & Err.Number & ", SQL = " & sSql
End Function

'----------------------------------------------------------------------------------------'
Public Function WritePswdPolicy(oSecCon As ADODB.Connection, sMessageParameters As String, sErrorMsg As String) As Boolean
'----------------------------------------------------------------------------------------'
'REM 12/11/02
'Writes the Password Policy to the Security database
'----------------------------------------------------------------------------------------'
Dim sSql As String
Dim rsPswdpolicy As ADODB.Recordset
Dim vPswdPolicy As Variant
Dim nMinLength As Integer
Dim nMaxLength As Integer
Dim nExpiryPeriod As Integer
Dim nEnforceMixedCase As Integer
Dim nEnforceDigit As Integer
Dim nAllowRrepeatChars As Integer
Dim nAllowUserName As Integer
Dim nPasswordHistory As Integer
Dim nPasswordRetries As Integer
    
    On Error GoTo ErrLabel

    vPswdPolicy = Split(sMessageParameters, gsPARAMSEPARATOR)
    
    nMinLength = vPswdPolicy(0)
    nMaxLength = vPswdPolicy(1)
    nExpiryPeriod = vPswdPolicy(2)
    nEnforceMixedCase = vPswdPolicy(3)
    nEnforceDigit = vPswdPolicy(4)
    nAllowRrepeatChars = vPswdPolicy(5)
    nAllowUserName = vPswdPolicy(6)
    nPasswordHistory = vPswdPolicy(7)
    nPasswordRetries = vPswdPolicy(8)

    'check to see if there is already a password policy in the database
    sSql = "SELECT COUNT (*) FROM MACROPassword"
    Set rsPswdpolicy = New ADODB.Recordset
    rsPswdpolicy.Open sSql, oSecCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    'if not then insert one
    If rsPswdpolicy.Fields(0) = 0 Then
        
        sSql = "INSERT INTO MACROPassword " _
            & " VALUES (" & nMinLength & ", " & nMaxLength & ", " & nExpiryPeriod & ", " & nEnforceMixedCase _
            & ", " & nEnforceDigit & ", " & nAllowRrepeatChars & ", " & nAllowUserName & ", " & nPasswordHistory & ", " & nPasswordRetries
        oSecCon.Execute sSql, , adCmdText
        
    Else 'else update existing one
    
        sSql = "UPDATE MACROPassword SET " _
            & " MINLENGTH = " & nMinLength & "," _
            & " MAXLENGTH = " & nMaxLength & "," _
            & " EXPIRYPERIOD = " & nExpiryPeriod & "," _
            & " ENFORCEMIXEDCASE = " & nEnforceMixedCase & "," _
            & " ENFORCEDIGIT = " & nEnforceDigit & "," _
            & " ALLOWREPEATCHARS = " & nAllowRrepeatChars & "," _
            & " ALLOWUSERNAME = " & nAllowUserName & "," _
            & " PASSWORDHISTORY = " & nPasswordHistory & "," _
            & " PASSWORDRETRIES = " & nPasswordRetries
        oSecCon.Execute sSql, , adCmdText
        
    End If
    
    rsPswdpolicy.Close
    Set rsPswdpolicy = Nothing
    
    WritePswdPolicy = True

Exit Function
ErrLabel:
    WritePswdPolicy = False
    sErrorMsg = "Error while writing new Password Policy. " & "Error Description: " & Err.Description _
                & "Error Number: " & Err.Number & ", SQL = " & sSql
End Function

'----------------------------------------------------------------------------------------'
Public Function WriteMsgToMessageTable(conMACRO As ADODB.Connection, sTrialSite As String, lClinicalTrialId As Long, nMessageType As Integer, _
                                  sMessageTimeStamp As String, sUserName As String, sMessageBody As String, sMessageParameters As String, _
                                  nMessageDirection As Integer, lOriginalMessageId As Long, nMessageTimeStamp_TZ As Integer, _
                                  nMessageReceived As Integer, ByRef sErrMessage As String) As String
'----------------------------------------------------------------------------------------'
'REM 12/11/02
'Writes a recieved message to the MACRO DB Message table and returns the messageid, received time stamp and time-zone offset
' REM 03/12/03 - Added ReplaceQuotes around sMessageParameters in the Insert statment
'----------------------------------------------------------------------------------------'
Dim sSql As String
Dim oTimezone As TimeZone
Dim nTimeZoneOffSet As Integer
Dim sMessageReceivedTimeStamp As String
Dim lMessageId As Long

    On Error GoTo ErrLabel

    Set oTimezone = New TimeZone
    
    'get the timestamp and the time-zone offset for local machine
    sMessageReceivedTimeStamp = SQLStandardNow
    nTimeZoneOffSet = oTimezone.TimezoneOffset
    
    'get next message id
    lMessageId = NextMessageId(conMACRO)
    
    'Insert the received message into MACRO DB Message table
    sSql = "INSERT INTO Message (TrialSite, ClinicalTrialId, MessageType, MessageTimeStamp, UserName, MessageBody," _
        & " MessageParameters, MessageReceived, MessageDirection, MessageId, MessageReceivedTimeStamp, MessageTimeStamp_TZ, MessageReceivedTimeStamp_TZ)" _
        & "  VALUES ('" & sTrialSite & "'," & lClinicalTrialId & "," & nMessageType & "," & sMessageTimeStamp & ",'" & sUserName & "','" & sMessageBody & "','" _
        & ReplaceQuotes(sMessageParameters) & "'," & nMessageReceived & "," & nMessageDirection & "," & lMessageId & "," _
        & sMessageReceivedTimeStamp & "," & nMessageTimeStamp_TZ & "," & nTimeZoneOffSet & ")"
    conMACRO.Execute sSql, , adCmdText
    
    'return a string contain the messageid, received time stamp and time-zone offset
    WriteMsgToMessageTable = lOriginalMessageId & gsPARAMSEPARATOR & sMessageReceivedTimeStamp & gsPARAMSEPARATOR & nTimeZoneOffSet
    
Exit Function
ErrLabel:
    sErrMessage = "Error while writing message to Message table. " & "Error Description: " & Err.Description _
                & ", Error Number: " & Err.Number & ", SQL = " & sSql
    WriteMsgToMessageTable = ""
End Function

'---------------------------------------------------------------------
Public Sub GetLogDetailsMessages(conMACRO As ADODB.Connection, sUserName As String)
'---------------------------------------------------------------------
'REM 19/11/02
'Adds all LogDetails to the MACRO database Message table so they are ready to collect and send
'REVISIONS:
' REM 17/03/04 - In sMessageParameters string, the vLogMsg(3,i) parameter now has a Replace around it to replace any "|" or "*" characters with the words "PIPE" or "STAR"
'---------------------------------------------------------------------
Dim vLogMsg As Variant
Dim sMessageParamaters As String
Dim nCount As Integer
Dim sConfirm As String
Dim sSiteCode As String
Dim i As Integer
Dim sLogDateTime As String
Dim sNextParameter As String

    On Error GoTo ErrLabel

    'returns site code from database if its a site dstatbase, or will return string = "server"
    sSiteCode = GetDBSettings(conMACRO, "datatransfer", "dbsitename", gsSERVER)
    
    'check to see if should insert LogDetails message into Message table, i.e. if server don't
    If InsertMessage(ExchangeMessageType.SystemLog, sSiteCode) Then
        'get all LogDetails entries that have a status on not sent (0)
        vLogMsg = LogDetails(conMACRO)
        
        If Not IsNull(vLogMsg) Then
            'set count to 0
            nCount = 0
            
            sMessageParamaters = ""
            sNextParameter = ""
            sConfirm = ""
            
            'loop through LogDetails
            For i = 0 To UBound(vLogMsg, 2)
                nCount = nCount + 1
                'REM 24/03/04 - convert to standard format
                sLogDateTime = LocalNumToStandard(vLogMsg(0, i))
                
                sNextParameter = sLogDateTime & gsPARAMSEPARATOR & vLogMsg(1, i) & gsPARAMSEPARATOR & vLogMsg(2, i) & gsPARAMSEPARATOR & Replace(Replace(vLogMsg(3, i), gsSEPARATOR, "PIPE"), gsPARAMSEPARATOR, "STAR") & gsPARAMSEPARATOR & vLogMsg(4, i) & gsPARAMSEPARATOR & vLogMsg(5, i) & gsPARAMSEPARATOR & sSiteCode
                
                'check to see if there are going to  be more than 4000 chars if you add the next parameter,
                'if so don't add next parameter
                If Len(sMessageParamaters & sNextParameter) < 4000 Then
                    sMessageParamaters = sMessageParamaters & sNextParameter
                    sConfirm = sConfirm & sLogDateTime
                    'only add the message separator if not the last message in the group of 10 or if not last message in array
                    If nCount <> 20 Then
                        If i <> UBound(vLogMsg, 2) Then
                            sMessageParamaters = sMessageParamaters & gsSEPARATOR
                            sConfirm = sConfirm & gsSEPARATOR
                        End If
                    End If
                Else
                    'If string was going to exceed 4000 chars then set counter to 20 and remove last delimiter off strings
                    nCount = 20
                    sMessageParamaters = Left(sMessageParamaters, (Len(sMessageParamaters) - 1))
                    sConfirm = Left(sConfirm, (Len(sConfirm) - 1))
                End If
                
                'adds the message to the message table in groups of 20 or if reaches last message in array
                'or if message got near to 4000 chars (then counter is set to 20 to stop string getting too long for DB)
                If (nCount = 20) Or (i = UBound(vLogMsg, 2)) Then
                    'add LoginLog messages to message table, then set LoginLog messages status to sent
                    If AddSystemMessage(conMACRO, ExchangeMessageType.SystemLog, sUserName, sUserName, "System Log", sMessageParamaters, "") Then
                        Call SetLogDetailsStatus(conMACRO, sConfirm)
                    End If
                    're-set counter to 0, sConfirm and sMessageParameters to ""
                    nCount = 0
                    sConfirm = ""
                    sMessageParamaters = ""
                End If
            Next
    
        End If
    
    End If

Exit Sub
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "modSysDataXfer.GetLogDetailsMessages"
End Sub

'---------------------------------------------------------------------
Public Sub GetUserLogMessages(oSecCon As ADODB.Connection, conMACRO As ADODB.Connection, sUserName As String)
'---------------------------------------------------------------------
'REM 15/11/02
'Adds all user log details to the MACRO database Message table ready to be collected and send
'adds them in batches of 20 Log messages per message
'REVISION:
'REM 23/03/04 - Added LocalNumToStandard around parameter in sMessageParameters as it is a date
'---------------------------------------------------------------------
Dim vLogMsg As Variant
Dim sMessageParamaters As String
Dim nCount As Integer
Dim sConfirm As String
Dim sSiteCode As String
Dim i As Integer
Dim sLogDateTime As String

    On Error GoTo ErrLabel
    
    'returns site code from database if its a site dstatbase, or will return string = "server"
    sSiteCode = GetDBSettings(conMACRO, "datatransfer", "dbsitename", gsSERVER)
    
    'check to see if should insert LoginLog message into Message table, i.e. if server don't
    If InsertMessage(ExchangeMessageType.UserLog, sSiteCode) Then
        vLogMsg = UserLogMessages(oSecCon)
        
        'check if there are any messages
        If Not IsNull(vLogMsg) Then
            'set count to 0
            nCount = 0
            
            sMessageParamaters = ""
            sConfirm = ""
            
            'loop through UserLog messages
            For i = 0 To UBound(vLogMsg, 2)
                nCount = nCount + 1
                'REM 23/03/04 - convert LocalNumToStandard around first parameter as it is a date
                sLogDateTime = LocalNumToStandard(vLogMsg(0, i))
                sMessageParamaters = sMessageParamaters & sLogDateTime & gsPARAMSEPARATOR & vLogMsg(1, i) & gsPARAMSEPARATOR & vLogMsg(2, i) & gsPARAMSEPARATOR & vLogMsg(3, i) & gsPARAMSEPARATOR & vLogMsg(4, i) & gsPARAMSEPARATOR & vLogMsg(5, i) & gsPARAMSEPARATOR & vLogMsg(6, i) & gsPARAMSEPARATOR & vLogMsg(7, i)
                sConfirm = sConfirm & sLogDateTime
                'only add the message separator if not the last message in the group of 10 or if not last message in array
                If nCount <> 20 Then
                    If i <> UBound(vLogMsg, 2) Then
                        sMessageParamaters = sMessageParamaters & gsSEPARATOR
                        sConfirm = sConfirm & gsSEPARATOR
                    End If
                End If
                
                'adds the message to the message table in groups of 20 or if reaches last message in array
                If (nCount = 20) Or (i = UBound(vLogMsg, 2)) Then
                    'add LoginLog messages to message table, then set LoginLog messages status to sent
                    If AddSystemMessage(conMACRO, ExchangeMessageType.UserLog, sUserName, sUserName, "User Log", sMessageParamaters, "") Then
                        Call SetLoginLogStatus(oSecCon, sConfirm)
                    End If
                    're-set counter to 0, sConfirm and sMessageParameters to ""
                    nCount = 0
                    sConfirm = ""
                    sMessageParamaters = ""
                End If
            Next
        End If
    End If

Exit Sub
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "modSysDataXfer.GetUserLogMessages"
End Sub

'---------------------------------------------------------------------
Private Function LogDetails(conMACRO As ADODB.Connection) As Variant
'---------------------------------------------------------------------
'REM 19/11/02
'returns all the LogDetails for a MACRO database
'---------------------------------------------------------------------
Dim sSql As String
Dim rsLogDetails As ADODB.Recordset

    On Error GoTo ErrLabel

    sSql = "SELECT * FROM LogDetails" _
        & " WHERE Status = 0" _
        & " ORDER BY LogDateTime"
    Set rsLogDetails = New ADODB.Recordset
    rsLogDetails.Open sSql, conMACRO, adOpenKeyset, adLockReadOnly, adCmdText

    If rsLogDetails.RecordCount > 0 Then
        LogDetails = rsLogDetails.GetRows
    Else
        LogDetails = Null
    End If
    
    rsLogDetails.Close
    Set rsLogDetails = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "modSysDataXfer.LogDetails"
End Function

'---------------------------------------------------------------------
Private Function UserLogMessages(oSecCon As ADODB.Connection) As Variant
'---------------------------------------------------------------------
'REM 19/11/02
'Returns array of all LoginLog entries that have not been sent to the server
'---------------------------------------------------------------------
Dim sSql As String
Dim rsLogMessages As ADODB.Recordset

    On Error GoTo ErrLabel

    'get all the login log records
    sSql = "SELECT * FROM LoginLog" _
        & " WHERE Status = 0" _
        & " ORDER BY LogDateTime"
    Set rsLogMessages = New ADODB.Recordset
    rsLogMessages.Open sSql, oSecCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    'check if there are any messages
    If rsLogMessages.RecordCount > 0 Then
        UserLogMessages = rsLogMessages.GetRows
    Else
        UserLogMessages = Null
    End If

    rsLogMessages.Close
    Set rsLogMessages = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "modSysDataXfer.UserLogMessages"
End Function

'---------------------------------------------------------------------
Private Sub SetLogDetailsStatus(conMACRO As ADODB.Connection, sConfirm As String)
'---------------------------------------------------------------------
'REM 19/11/02
'Sets the logDetails status to sent (1)
'---------------------------------------------------------------------
Dim sSql As String
Dim vLogDatetime As Variant
Dim dLogDateTime As String
Dim i As Integer
    
    On Error GoTo ErrLabel

    vLogDatetime = Split(sConfirm, gsSEPARATOR)
    
    For i = 0 To UBound(vLogDatetime)
        'NB - vLogDateTime is in standard format
        sSql = "UPDATE LogDetails SET Status = 1" _
            & " WHERE Status = 0" _
            & " AND LogDateTime = " & vLogDatetime(i)
        conMACRO.Execute sSql
    
    Next
    
Exit Sub
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "modSysDataXfer.SetLogDetailsStatus"
End Sub

'---------------------------------------------------------------------
Private Sub SetLoginLogStatus(oSecCon As ADODB.Connection, sConfirm As String)
'---------------------------------------------------------------------
'REM 18/11/02
'Set the LoginLog status
'---------------------------------------------------------------------
Dim sSql As String
Dim vLogDatetime As Variant
Dim dLogDateTime As String
Dim i As Integer
    
    On Error GoTo ErrLabel

    vLogDatetime = Split(sConfirm, gsSEPARATOR)
    
    For i = 0 To UBound(vLogDatetime)
        
        sSql = "UPDATE LoginLog SET Status = 1" _
            & " WHERE Status = 0" _
            & " AND LogDateTime = " & vLogDatetime(i)
        oSecCon.Execute sSql
    
    Next
    
Exit Sub
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "modSysDataXfer.SetLoginLogStatus"
End Sub

'---------------------------------------------------------------------
Public Function RestoreUserRoles(conMACRO As ADODB.Connection, sSiteCode As String) As Variant
'---------------------------------------------------------------------
'REM 28/11/02
'Gets all the restore site UserRole messages from the server message table (there are never any of these messages on a site database)
'---------------------------------------------------------------------
Dim sSql As String
Dim rsRestore As ADODB.Recordset

    On Error GoTo ErrLabel
    
    sSql = "SELECT * FROM Message" _
        & " WHERE MessageType = " & ExchangeMessageType.RestoreUserRole _
        & " AND MessageReceived = " & MessageReceived.NotYetReceived _
        & " AND TrialSite = '" & sSiteCode & "'" _
        & " Order BY MessageTimeStamp DESC"
    
    Set rsRestore = New ADODB.Recordset
    rsRestore.Open sSql, conMACRO, adOpenKeyset, adLockReadOnly, adCmdText

    If rsRestore.RecordCount > 0 Then
        RestoreUserRoles = rsRestore.GetRows
    Else
        RestoreUserRoles = Null
    End If
    
    rsRestore.Close
    Set rsRestore = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "modSysDataXfer.RestoreUserRoles"
End Function

'---------------------------------------------------------------------
Public Function SystemMessages(conMACRO As ADODB.Connection, sSiteCode As String) As Variant
'---------------------------------------------------------------------
'REM 13/11/02
'Returns all system messages that have not been sent
'---------------------------------------------------------------------
Dim sSql As String
Dim rsMessages As ADODB.Recordset

    On Error GoTo ErrLabel
    
    sSql = "SELECT * FROM Message" _
        & " WHERE MessageType IN (" & gsSYSTEM_MESSAGE_TYPES & ")" _
        & " AND MessageReceived = " & MessageReceived.NotYetReceived
    If LCase(sSiteCode) <> "" Then
        sSql = sSql & " AND TrialSite = '" & sSiteCode & "'"
    End If
        sSql = sSql & "ORDER BY MessageTimeStamp"
    
    Set rsMessages = New ADODB.Recordset
    rsMessages.Open sSql, conMACRO, adOpenKeyset, adLockReadOnly, adCmdText

    If rsMessages.RecordCount > 0 Then
        SystemMessages = rsMessages.GetRows
    Else
        SystemMessages = Null
    End If
    
    rsMessages.Close
    Set rsMessages = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "modSysDataXfer.GetSystemMessages"
End Function

'---------------------------------------------------------------------
Public Function UserPasswordAndDetails(conMACRO As ADODB.Connection, sSiteCode As String) As Variant
'---------------------------------------------------------------------
'REM 28/11/02
'Returns all the Change Password and User Details messages
'---------------------------------------------------------------------
Dim sSql As String
Dim rsPswds As ADODB.Recordset

    On Error GoTo ErrLabel

    sSql = "SELECT * FROM Message" _
        & " WHERE MessageType IN (" & ExchangeMessageType.PasswordChange & "," & ExchangeMessageType.User & ")" _
        & " AND MessageReceived = " & MessageReceived.NotYetReceived _
        & " AND MessageDirection = " & MessageDirection.MessageOut _
        & " AND TrialSite = '" & sSiteCode & "'"
    Set rsPswds = New ADODB.Recordset
    rsPswds.Open sSql, conMACRO, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsPswds.RecordCount <> 0 Then
        UserPasswordAndDetails = rsPswds.GetRows
    Else
        UserPasswordAndDetails = Null
    End If
    
    rsPswds.Close
    Set rsPswds = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "modSysDataXfer.UserPasswordAndDetails"
End Function

'---------------------------------------------------------------------
Public Sub ConfirmMessages(conMACRO As ADODB.Connection, sConfirmMessages As String, sErrorMsg As String)
'---------------------------------------------------------------------
'REM 13/11/02
'Changes status of passed in message Id's to sent
'---------------------------------------------------------------------
Dim sSql As String
Dim vMessages As Variant
Dim sMessage As String
Dim vMessageParameters As Variant
Dim lMessageId As Long
Dim sMessageReceivedTimeStamp As String
Dim nTimeZoneOffSet As Integer
Dim i As Integer
Dim j As Integer

    On Error GoTo ErrLabel
    
    'get all messages from the passed in string
    vMessages = Split(sConfirmMessages, gsSEPARATOR)
    'loop through the message, changing each one to recieved and setting the received timestamp and time zone offset
    For i = 0 To UBound(vMessages)
        
        sMessage = vMessages(i)
         
        'get the parameters for each message
        vMessageParameters = Split(sMessage, gsPARAMSEPARATOR)
        
        lMessageId = vMessageParameters(0)
        sMessageReceivedTimeStamp = vMessageParameters(1)
        nTimeZoneOffSet = vMessageParameters(2)
        
        'update the messages in the message table to received
        sSql = "UPDATE Message SET MessageReceived = " & MessageReceived.Received & ", " _
            & " MessageReceivedTimeStamp = " & sMessageReceivedTimeStamp & ", " _
            & " MessageReceivedTimeStamp_TZ = " & nTimeZoneOffSet _
            & " WHERE MessageId = " & lMessageId
        conMACRO.Execute sSql

     Next
    
Exit Sub
ErrLabel:
    sErrorMsg = "Error in confirming sent messages. Error Description: " & Err.Description & ", Error Number: " _
                    & Err.Number & ", SQL = " & sSql
End Sub

'----------------------------------------------------------------------------------------'
Public Function CreateConnection(oSecCon As ADODB.Connection, sDatabaseCode As String, ByRef sErrorMsg As String, Optional ByRef sConnString As String) As ADODB.Connection
'----------------------------------------------------------------------------------------'
'REM 11/11/02
'Create a connection for a given database
'----------------------------------------------------------------------------------------'
' REVISIONS
' DPH 28/01/2003 - Return connection string
'----------------------------------------------------------------------------------------'
Dim conMACRO As ADODB.Connection
Dim sMessage As String
Dim sConnection As String

    On Error GoTo ErrLabel
    
    sConnection = ConnectionString(sDatabaseCode, sMessage, oSecCon)
        
    If sConnection <> "" Then
        
        'create connection for selected database
        Set conMACRO = New ADODB.Connection
    
        conMACRO.Open sConnection
        conMACRO.CursorLocation = adUseClient
        
        Set CreateConnection = conMACRO
    
    End If
    
    sConnString = sConnection
    
Exit Function
ErrLabel:
    sErrorMsg = "MACRO Database connection error: " & Err.Description & ": Error no. " & Err.Number
End Function

'---------------------------------------------------------------------
Public Function ConnectionString(sDatabaseCode As String, ByRef sErrorMsg As String, Optional oSecCon As ADODB.Connection = Nothing) As String
'---------------------------------------------------------------------
'REM 26/11/02
'Create a MACRO database connection string for a given database
'---------------------------------------------------------------------
Dim oDatabase As MACROUserBS30.Database
Dim sConnection As String

    On Error GoTo ErrLabel

    Set oDatabase = New MACROUserBS30.Database
    
    If oSecCon Is Nothing Then
        Set oSecCon = CreateSecurityConnection(sErrorMsg)
    End If

    If oDatabase.Load(oSecCon, "", sDatabaseCode, "", False, sErrorMsg) Then
    
        ConnectionString = oDatabase.ConnectionString
    Else
        ConnectionString = sErrorMsg
    End If
    
    Set oDatabase = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "basSystemTransfer.ConnectionString"
End Function

'---------------------------------------------------------------------
Public Function GetDBSettings(conMACRO As ADODB.Connection, sSettingSection As String, sSettingKey As String, sDefault As String) As String
'---------------------------------------------------------------------
'REM 14/11/02
'Returns whether a database is a Serever of Site DB
'---------------------------------------------------------------------
Dim sSql As String
Dim rsDBSetting As ADODB.Recordset

    On Error GoTo ErrLabel

    sSql = "SELECT SettingValue FROM MACRODBSetting" _
        & " WHERE SettingSection = '" & sSettingSection & "'" _
        & " AND SettingKey = '" & sSettingKey & "'"
     Set rsDBSetting = New ADODB.Recordset
     rsDBSetting.Open sSql, conMACRO, adOpenKeyset, adLockPessimistic, adCmdText

    If rsDBSetting.RecordCount <> 0 Then
        GetDBSettings = rsDBSetting![SettingValue]
    Else
        GetDBSettings = sDefault
    End If

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "modSysDataXfer.GetDBSettings"
End Function

'---------------------------------------------------------------------
Public Function InsertMessage(nMessageType As Integer, sTrialSite As String) As Boolean
'---------------------------------------------------------------------
'REM 15/11/02
'Checks to see if a given message type from a site or server should be inseretd
'into the message table ready for data transfer
'---------------------------------------------------------------------

    On Error GoTo ErrLabel

    If LCase(sTrialSite) = gsSERVER Then
        Select Case nMessageType
        Case 32, 33, 34, 35, 38, 40
            InsertMessage = True
        Case Else
            InsertMessage = False
        End Select
    Else
        Select Case nMessageType
        Case 32, 33, 34, 36, 37
            InsertMessage = True
        Case Else
            InsertMessage = False
        End Select
    End If
    
Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "modSysDataXfer.InsertMessage"
End Function

'----------------------------------------------------------------------------------------'
Public Function AddSystemMessage(conMACRO As ADODB.Connection, nMessageType As Integer, sSystemAdminUserName As String, _
                                 sUserName As String, sMessageBody As String, sMessageParameters As String, _
                                 sSite As String, Optional sRoleCode As String, Optional sMsgSite As String = "") As Boolean
'----------------------------------------------------------------------------------------'
'REM 14/11/02
'Adds a system message to the MACRO Message table
'Routine checks to see if the passed in message type should be written to the message table
'Certain message types are not written to the message table if the database is either a site or server
'sSite: the site that you want the message distributed to (empty string if want to distribute to all user sites)
'sMsgSite: The site an incoming message came from, used so you don't dristribute a message back to the site it came from
'----------------------------------------------------------------------------------------'
Dim sSql As String
Dim nMessageDirection As Integer
Dim vSites As Variant
Dim i As Integer
Dim sSiteCode As String

    On Error GoTo ErrLabel
    
    'get the trialsite name, will be "server" if finds nothing in MACRODBsetting table
    sSiteCode = GetDBSettings(conMACRO, "datatransfer", "dbsitename", gsSERVER)

    AddSystemMessage = False
    
    'check to see if the passed in message should be inserted into the message table
    If InsertMessage(nMessageType, sSiteCode) Then

        'message direction, i.e. site(1) or server(0)
        If LCase(sSiteCode) = gsSERVER Then
            nMessageDirection = MessageDirection.MessageOut ' 0 server to site
        Else
            nMessageDirection = MessageDirection.MessageIn '1 site to server
        End If
        
        Select Case nMessageType
        
        Case ExchangeMessageType.User, ExchangeMessageType.PasswordChange, ExchangeMessageType.SystemLog, ExchangeMessageType.UserLog ' 32, 34, 36, 37 'User, Password change, system log, user log
            'if database is a server DB then get sites that message must be distributed to
            If LCase(sSiteCode) = gsSERVER Then
                If sSite = "" Then
                    'get user sites
                    vSites = UserSites(conMACRO, sUserName, sRoleCode)
                
                    'If the user has been assigned sites, then loop through them
                    If Not IsNull(vSites) Then
                        For i = 0 To UBound(vSites, 2)
                            sSiteCode = vSites(0, i)
                            'make sure message is not disributed back to site it came from (only used when site message is received by server and then distributed to other relevant sites)
                            If sSiteCode <> sMsgSite Then
                                Call InsertNewMessage(conMACRO, sSiteCode, nMessageType, sSystemAdminUserName, sMessageBody, sMessageParameters, nMessageDirection)
                            End If
                        Next
                    End If
                    
                Else 'distribute to a specified site
                    sSiteCode = sSite
                    Call InsertNewMessage(conMACRO, sSiteCode, nMessageType, sSystemAdminUserName, sMessageBody, sMessageParameters, nMessageDirection)

                End If
                
            Else 'else if its a site, need only one message to be sent to the server
                Call InsertNewMessage(conMACRO, sSiteCode, nMessageType, sSystemAdminUserName, sMessageBody, sMessageParameters, nMessageDirection)
                
            End If
            
        Case ExchangeMessageType.UserRole, ExchangeMessageType.RestoreUserRole ' 33, 38 'new user role, Restore userrole
                If LCase(sSiteCode) = gsSERVER Then
                    'check to see if UserRole is for all sites, if so then send message to all sites
                    If sSite = "AllSites" Then
                        vSites = AllSites(conMACRO)
                        If Not IsNull(vSites) Then
                            For i = 0 To UBound(vSites, 2)
                                sSiteCode = vSites(0, i)
                                'make sure message is not disributed back to site it came from (only used when site message is received and then distributed to other relevant sites)
                                If sSiteCode <> sMsgSite Then
                                    Call InsertNewMessage(conMACRO, sSiteCode, nMessageType, sSystemAdminUserName, sMessageBody, sMessageParameters, nMessageDirection)
                                End If
                            Next
                        End If
                    Else 'else is for specific site
                            sSiteCode = sSite
                            'make sure you do not distribute message back to site it just came from
                            If sSiteCode <> sMsgSite Then
                                Call InsertNewMessage(conMACRO, sSiteCode, nMessageType, sSystemAdminUserName, sMessageBody, sMessageParameters, nMessageDirection)
                            End If
                    End If
                Else 'its a site only send one message
                    Call InsertNewMessage(conMACRO, sSiteCode, nMessageType, sSystemAdminUserName, sMessageBody, sMessageParameters, nMessageDirection)
                End If
        
        Case ExchangeMessageType.Role, ExchangeMessageType.PasswordPolicy ' 35, 40 'New/updated role and Password policy
            If sSite = "" Then
                'these will be distributed to all sites
                vSites = AllSites(conMACRO)
                If Not IsNull(vSites) Then
                    For i = 0 To UBound(vSites, 2)
                        sSiteCode = vSites(0, i)
                        Call InsertNewMessage(conMACRO, sSiteCode, nMessageType, sSystemAdminUserName, sMessageBody, sMessageParameters, nMessageDirection)
                    Next
                End If
            Else 'distribute to a specific site
                sSiteCode = sSite
                Call InsertNewMessage(conMACRO, sSiteCode, nMessageType, sSystemAdminUserName, sMessageBody, sMessageParameters, nMessageDirection)
            End If
        
        End Select
        AddSystemMessage = True

    End If
    
Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "modSysDataXfer.AddSystemMessage"

End Function

'---------------------------------------------------------------------
Private Sub InsertNewMessage(conMACRO As ADODB.Connection, sSiteCode As String, nMessageType As Integer, _
                          sSystemAdminUserName As String, sMessageBody As String, _
                          sMessageParameters As String, nMessageDirection As Integer)
'---------------------------------------------------------------------
'REM 20/11/02
'Inserts a message into the message table
'---------------------------------------------------------------------
Dim sSql As String
Dim lClinicalTrialId As Long
Dim oTimezone As TimeZone
Dim nTimeZoneOffSet As Integer
Dim sMessageTimeStamp As String
Dim lMessageId As Long
    
    On Error GoTo ErrLabel
    
    Set oTimezone = New TimeZone

    'get the timestamp and the time-zone offset for local machine
    sMessageTimeStamp = SQLStandardNow
    
    nTimeZoneOffSet = oTimezone.TimezoneOffset

    'get next message id
    lMessageId = NextMessageId(conMACRO)
    lClinicalTrialId = -1
    
    'Insert message into MACRO DB Message table
    sSql = "INSERT INTO Message (TrialSite, ClinicalTrialId, MessageType, MessageTimeStamp, UserName, MessageBody," _
        & " MessageParameters, MessageReceived, MessageDirection, MessageId, MessageTimeStamp_TZ)" _
        & "  VALUES ('" & sSiteCode & "'," & lClinicalTrialId & "," & nMessageType & "," & sMessageTimeStamp & ",'" _
        & sSystemAdminUserName & "','" & sMessageBody & "','" & ReplaceQuotes(sMessageParameters) & "'," & MessageReceived.NotYetReceived & "," _
        & nMessageDirection & "," & lMessageId & "," & nTimeZoneOffSet & ")"
    conMACRO.Execute sSql, , adCmdText
    
Exit Sub
ErrLabel:
Err.Raise Err.Number, , Err.Description & ", SQL = " & sSql & "|" & "modSysDataXfer.InsertNewMessage"
End Sub

'---------------------------------------------------------------------
Private Function NextMessageId(conMACRO As ADODB.Connection) As Long
'---------------------------------------------------------------------
'REM 19/11/02
'Returns the next Message Id
'   TA  18/01/2006 - MessageId now calculated by a sequence to avoid duplicate id problem
'---------------------------------------------------------------------
Dim xfer As SysDataXfer

    Set xfer = New SysDataXfer
    NextMessageId = xfer.GetNextMessageId(conMACRO)
    Set xfer = Nothing
        
Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "modSysDataXfer.NextMessageId"
End Function

'---------------------------------------------------------------------
Private Function UserSites(conMACRO As ADODB.Connection, sUserName As String, sRoleCode As String) As Variant
'---------------------------------------------------------------------
'REM 19/11/02
'Returns all the sites a user has access to
'---------------------------------------------------------------------
Dim sSql As String
Dim rsSites As ADODB.Recordset
    
    On Error GoTo ErrLabel


    'see if the user has AllSites
    sSql = "SELECT SiteCode FROM UserRole WHERE SiteCode = 'AllSites'"
    
    'if there is no rolecode passed in then don't select on it
    If sRoleCode <> "" Then
        sSql = sSql & " AND RoleCode = '" & sRoleCode & "'"
    End If
    
    sSql = sSql & " AND UserName = '" & sUserName & "'"
    
    Set rsSites = New ADODB.Recordset
    rsSites.Open sSql, conMACRO, adOpenKeyset, adLockReadOnly, adCmdText

    'if not AllSites then get Sites user pass from UserRole table
    If rsSites.RecordCount = 0 Then
        sSql = "SELECT DISTINCT SiteCode FROM UserRole WHERE UserName = '" & sUserName & "'"
        'if there is no rolecode passed in then don't select on it
        If sRoleCode <> "" Then
            sSql = sSql & " AND RoleCode = '" & sRoleCode & "'"
        End If
        Set rsSites = New ADODB.Recordset
        rsSites.Open sSql, conMACRO, adOpenKeyset, adLockReadOnly, adCmdText
    Else 'get all sites
        sSql = "SELECT Site" _
            & " FROM Site"
        Set rsSites = New ADODB.Recordset
        rsSites.Open sSql, conMACRO, adOpenKeyset, adLockReadOnly, adCmdText
        
    End If

    If rsSites.RecordCount <> 0 Then
        UserSites = rsSites.GetRows
    Else
        UserSites = Null
    End If

    rsSites.Close
    Set rsSites = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "modSysDataXfer.UserSites"
End Function

'---------------------------------------------------------------------
Private Function AllSites(conMACRO As ADODB.Connection)
'---------------------------------------------------------------------
'REM 19/11/02
'Returns all the sites in a database
'---------------------------------------------------------------------
Dim rsSites As ADODB.Recordset
Dim sSql As String
    
    On Error GoTo ErrLabel

    sSql = "SELECT Site" _
        & " FROM Site"
    Set rsSites = New ADODB.Recordset
    rsSites.Open sSql, conMACRO, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsSites.RecordCount <> 0 Then
        AllSites = rsSites.GetRows
    Else
        AllSites = Null
    End If
    
    rsSites.Close
    Set rsSites = Nothing
    
Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "modSysDataXfer.AllSites"
End Function

'---------------------------------------------------------------------
Private Function CheckPendingMessages(conMACRO As ADODB.Connection, sTrialSite As String, sUserName As String, _
                                          nMessageType As Integer, sSiteCode As String) As Boolean
'---------------------------------------------------------------------
'REM 27/11/02
'Checks to see if there are any pending messages to be sent from the server for a specific user
'---------------------------------------------------------------------
Dim sSql As String
Dim rsUser As ADODB.Recordset
Dim vUser As Variant
Dim sUser As String
Dim vUserDetails As Variant
Dim vMessage As Variant
Dim sMessageParameters As String
Dim sPendingUserName As String
Dim i As Integer

    On Error GoTo ErrLabel
    
    If LCase(sSiteCode) = gsSERVER Then
    
        'get all user messages on the server for the site that have not been sent yet
        sSql = "SELECT * FROM Message" _
            & " WHERE MessageType = " & nMessageType _
            & " AND MessageReceived = 0" _
            & " AND TrialSite = '" & sTrialSite & "'"
        Set rsUser = New ADODB.Recordset
        rsUser.Open sSql, conMACRO, adOpenKeyset, adLockReadOnly, adCmdText
        
        If rsUser.RecordCount > 0 Then
            'go through each message
            vUser = rsUser.GetRows
            For i = 0 To UBound(vUser, 2)

                'get the parameters field
                sMessageParameters = vUser(6, i)
                'get user details
                vUserDetails = Split(sMessageParameters, gsPARAMSEPARATOR)
                'get user name from details
                sPendingUserName = vUserDetails(0)
                        
                'check to see if it matches the one passed in
                If LCase(sPendingUserName) = LCase(sUserName) Then
                    CheckPendingMessages = True
                    Exit Function
                Else
                    CheckPendingMessages = False
                End If
                
            Next
        Else
            CheckPendingMessages = False
        End If
        
        rsUser.Close
        Set rsUser = Nothing
    Else 'if its a site then return false as don't check for any pending messages
        CheckPendingMessages = False
    End If

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "modSysDataXfer.CheckPendingMessages"
End Function

'---------------------------------------------------------------------
Private Sub AddUserDatabase(oSecCon As ADODB.Connection, sUserName As String, sDatabaseCode As String)
'---------------------------------------------------------------------
'REM 29/11/02
'Checks to see if a User/Database combination exists in the UserDatabase table, if not adds it
'---------------------------------------------------------------------
Dim sSql As String
Dim rsUserDB As ADODB.Recordset

    On Error GoTo ErrHandler

    sSql = "SELECT COUNT (*) FROM UserDatabase" _
        & " WHERE UserName = '" & sUserName & "'" _
        & " AND DatabaseCode = '" & sDatabaseCode & "'"
    Set rsUserDB = New ADODB.Recordset
    rsUserDB.Open sSql, oSecCon, adOpenKeyset, adLockOptimistic, adCmdText
    
    If rsUserDB.Fields(0) = 0 Then
        sSql = "INSERT INTO UserDatabase " _
            & " VALUES ('" & sUserName & "','" & sDatabaseCode & "')"
            oSecCon.Execute sSql, , adCmdText
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modSysDataXfer.AddUserDatabase"
End Sub

'---------------------------------------------------------------------
Private Sub DeleteUserDatabase(conMACRO As ADODB.Connection, oSecCon As ADODB.Connection, sUserName As String, sDatabaseCode As String)
'---------------------------------------------------------------------
'REM 29/11/02
'Checks to see if there are any UserRoles for a given user in a database, if not removes the User/Database combination
'from the UserDatabase table in the Security database
'---------------------------------------------------------------------
Dim sSql As String
Dim rsUserDB As ADODB.Recordset

    On Error GoTo ErrHandler
    
    sSql = "SELECT COUNT (*) FROM UserRole" _
        & " WHERE UserName = '" & sUserName & "'"
    Set rsUserDB = New ADODB.Recordset
    rsUserDB.Open sSql, conMACRO, adOpenKeyset, adLockOptimistic, adCmdText

    If rsUserDB.Fields(0) = 0 Then
        sSql = "DELETE FROM UserDatabase" _
            & " WHERE UserName = '" & sUserName & "'" _
            & " AND DatabaseCode = '" & sDatabaseCode & "'"
        oSecCon.Execute sSql, , adCmdText
    End If
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modSysDataXfer.DeleteUserDatabase"
End Sub

'---------------------------------------------------------------------
Public Function SQLStandardNow() As String
'---------------------------------------------------------------------
' Returns Now as a double in STANDARD numeric format
'---------------------------------------------------------------------

    SQLStandardNow = LocalNumToStandard(IMedNow)

End Function

'---------------------------------------------------------------------
Public Function CreateReportFilesZIP(conMACRO As ADODB.Connection, oSecCon As ADODB.Connection, _
                                    sUserName As String, sDatabase As String, _
                                    sSiteCode As String, sLastTransDate As String, ByRef bNoFiles As Boolean, _
                                    sMACROConn As String) As String
'---------------------------------------------------------------------
' Get Report Files which have been modified since last site transfer
' Zip them up and place in published HTML folder
' Return the name of the created zip file
'---------------------------------------------------------------------
'REM 13/10/03 - Pass last transfer date into routine (sent from site), no longer get it from server Logdetails table
'---------------------------------------------------------------------
    Dim sReportFilePath As String
    Dim sPublishedHTMLPath As String
    Dim sFileName As String
    Dim rsUser As ADODB.Recordset
    Dim sSql As String
    Dim sFileList() As String
    Dim dateLastXfer As Date
    
    ' Get server file paths
    Call ReturnFilePaths(oSecCon, sDatabase, sPublishedHTMLPath, sReportFilePath)
    
    ' collect date/time of last successful transfer from site
    'REM 13/10/03 - Pass in date of last transfer (is sent from the site now)
    dateLastXfer = CDate(CDbl(StandardNumToLocal(sLastTransDate)))  'ReturnLastXFerDateTime(conMACRO, sSiteCode, sMACROConn)
    
    ' get file list of modified reports
    sFileList = ReturnReportFileList(dateLastXfer, sReportFilePath)
    
    ' if no filelist
    If sFileList(0) <> "" Then
        ' zip up report file list
        sFileName = ZipUpFileList(conMACRO, sUserName, sFileList, sPublishedHTMLPath, sSiteCode)
        bNoFiles = False
    Else
        sFileName = ""
        bNoFiles = True
    End If
        
    ' return filename
    CreateReportFilesZIP = sFileName
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modSysDataXfer.CreateReportFilesZIP"
End Function

'---------------------------------------------------------------------
Private Function ReturnLastXFerDateTime(conMACRO As ADODB.Connection, sSiteCode As String, _
                                        sMACROConn As String) As Date
'---------------------------------------------------------------------
' Get most recent date/time from LogDetails
'---------------------------------------------------------------------
Dim sSql As String
Dim rsMessageTime  As ADODB.Recordset
    
    On Error GoTo ErrHandler

    sSql = "SELECT DISTINCT LogDateTime FROM LogDetails WHERE TaskId = 'ReportXfer' " _
        & "AND "
        
    Select Case Connection_Property(CONNECTION_PROVIDER, sMACROConn)
    Case CONNECTION_MSDAORA, CONNECTION_ORAOLEDB_ORACLE
        sSql = sSql & " (NLS_LOWER(LogMessage) LIKE '%" _
                        & Replace(LCase("Site " & sSiteCode), "'", "''") & "%') "
    Case Else
        sSql = sSql & " LogMessage Like '%" & Replace(LCase("Site " & sSiteCode), "'", "''") & "%' "
    End Select
    
    sSql = sSql & " ORDER BY LogDateTime DESC"
    
    Set rsMessageTime = New ADODB.Recordset
    rsMessageTime.Open sSql, conMACRO, adOpenForwardOnly, , adCmdText
    
    ' Get first record
    If Not rsMessageTime.EOF Then
        ReturnLastXFerDateTime = CDate(rsMessageTime.Fields("LogDateTime"))
    Else
        ReturnLastXFerDateTime = "01/01/1980 00:00"
    End If
    
    rsMessageTime.Close
    Set rsMessageTime = Nothing
     
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modSysDataXfer.ReturnLastXFerDateTime"
End Function

'---------------------------------------------------------------------
Private Sub ReturnFilePaths(conSecMACRO As ADODB.Connection, ByVal sDatabaseCode As String, _
                            ByRef sHTMLPath As String, _
                            ByRef sReportPath As String)
'---------------------------------------------------------------------
' Get File paths from security database
' TODO - Reports Setting - Hard coded for moment
'---------------------------------------------------------------------
Dim sSql As String
Dim rsPaths As ADODB.Recordset
    
    On Error GoTo ErrHandler

    sSql = "SELECT HTMLLOCATION,REPORTSLOCATION FROM DATABASES WHERE DATABASECODE = '" & sDatabaseCode & "'"
        
    Set rsPaths = New ADODB.Recordset
    rsPaths.Open sSql, conSecMACRO, adOpenForwardOnly, , adCmdText
    
    ' Get Paths
    If Not rsPaths.EOF Then
        If IsNull(rsPaths.Fields("HTMLLOCATION")) Then
            sHTMLPath = ""
        Else
            sHTMLPath = rsPaths.Fields("HTMLLOCATION")
        End If
        If IsNull(rsPaths.Fields("REPORTSLOCATION")) Then
            sReportPath = ""
        Else
            sReportPath = rsPaths.Fields("REPORTSLOCATION")
        End If
'        sReportPath = "F:\TrialOfficeStorage\Reports\"
    Else
        sHTMLPath = ""
        sReportPath = ""
    End If
    
    ' if no path then default
    If sHTMLPath = "" Then
        sHTMLPath = App.Path & "\..\html\"
    End If
    If sReportPath = "" Then
        sReportPath = App.Path & "\..\www\reports\"
    End If
    
    rsPaths.Close
    Set rsPaths = Nothing
     
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modSysDataXfer.ReturnFilePaths"
End Sub

'---------------------------------------------------------------------
Private Function ReturnReportFileList(dateLastXfer As Date, sReportPath As String) As String()
'---------------------------------------------------------------------
' Get the filelist from the report folder with a new modified date
'---------------------------------------------------------------------
Dim sFileList() As String
Dim lArray As Long
Dim oFSO As New FileSystemObject
Dim oFSFolder As Folder
Dim colFiles As Files
Dim oFSFile As File
Dim nFileNo As Integer
Dim sFileAndPath As String

    On Error GoTo ErrHandler
    
    ReDim sFileList(0)
    
    If oFSO.FolderExists(sReportPath) Then
        ' Get collection of files in the folder from a folder object
        Set oFSFolder = oFSO.GetFolder(sReportPath)
        Set colFiles = oFSFolder.Files
        ' loop through files and find recently modified ones
        For Each oFSFile In colFiles
        'For nFileNo = 1 To colFiles.Count
            'Set oFSFile = colFiles(nFileNo)
            ' is it a recently modified file
            If oFSFile.DateLastModified > dateLastXfer Then
                sFileAndPath = oFSFile.Path
                ' add to filelist array
                lArray = UBound(sFileList)
                If lArray = 0 And sFileList(lArray) = "" Then
                    sFileList(lArray) = sFileAndPath
                Else
                    ReDim Preserve sFileList(lArray + 1)
                    sFileList(lArray + 1) = sFileAndPath
                End If
            End If
        Next
    Else
        ' folder doesn't exist
        sFileList(0) = ""
    End If
    
    Set oFSO = Nothing
    
    ReturnReportFileList = sFileList
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modSysDataXfer.ReturnReportFileList"
End Function

'---------------------------------------------------------------------
Private Function ZipUpFileList(conMACRO As ADODB.Connection, sUserName As String, _
                                sFileList() As String, sHTMLPath As String, _
                                sSiteCode As String) As String
'---------------------------------------------------------------------
' Zip up the appropriate filelist and write to published HTML folder
'---------------------------------------------------------------------
Dim sZipFileName As String
Dim sZipPathAndFileName As String

On Error GoTo ErrHandler
    ' Zip file name is site_study_"Reports"_datetime.zip
    sZipFileName = sSiteCode & "_Reports_" & _
                Format(CDate(IMedNow), "YYYYMMDDHHmmss") & ".zip"
         
    sZipPathAndFileName = sHTMLPath & sZipFileName
    
    Call ZipFiles(sFileList, sZipPathAndFileName)
    
    ZipUpFileList = sZipFileName
Exit Function
ErrHandler:
    Call gLogForXfer("ReportXferErr", "Site " & sSiteCode & " failed to create ZIP file. Err No: " & Err.Number & _
            " Desc: " & Err.Description, conMACRO, sUserName)
    ZipUpFileList = ""
End Function

'---------------------------------------------------------------------
Public Function StoreReportFiles(conMACRO As ADODB.Connection, oSecCon As ADODB.Connection, _
                                    sDatabase As String, sSiteCode As String, _
                                    sReportZipFile As String) As Boolean
'---------------------------------------------------------------------
' Unzip file containing report files
' Write to local Report file folder
'---------------------------------------------------------------------
    Dim sReportFilePath As String
    Dim sPublishedHTMLPath As String
    
    ' Get file paths (but don't need to use HTML path here)
    Call ReturnFilePaths(oSecCon, sDatabase, sPublishedHTMLPath, sReportFilePath)
    
    ' extract files from zip to reports folder
    Call UnZipFiles(sReportFilePath, sReportZipFile)
    
    ' return filename
    StoreReportFiles = True
    
Exit Function
ErrHandler:
    StoreReportFiles = False
    Err.Raise Err.Number, , Err.Description & "|modSysDataXfer.StoreReportFiles"
End Function

'---------------------------------------------------------------------
Public Sub gLogForXfer(ByVal sTaskId As String, ByVal sMessage As String, _
                ByRef sConMACRO As ADODB.Connection, ByVal sUserName As String)
'---------------------------------------------------------------------
' Added Writing to system Log function
'---------------------------------------------------------------------
Dim sSql As String
Dim sSQLNow As String
Dim nLogNumber As Long
Dim rsLogDetails As ADODB.Recordset
Dim nTimeZone As Integer
Dim sLocation As String
Dim oTimezone As TimeZone
'Ash 12/12/2002
Dim conNewMACRO As ADODB.Connection

    On Error GoTo ErrHandler
    
    Set conNewMACRO = sConMACRO
    
    Set oTimezone = New TimeZone
    
    ' NCJ 4 Feb 00 SR2851 - Use standard SQL datestamp
    sSQLNow = SQLStandardNow
    
    'Log messages have a combined key of LogDateTime and LogNumber. The first log
    'messages for a particular time will have a LogNumber of 0, the next 1 and so on
    'until the LogDateTime moves on a second.
    
    ' NCJ 4/2/00 - Changed dNow to sSQLNow
    sSql = " SELECT LogNumber From LogDetails WHERE LogDateTime = " & sSQLNow

    'assess the number of records and set the LogNumber for this entry (nLogNumber)
    Set rsLogDetails = New ADODB.Recordset
    'Note use of adOpenKeyset cusor. Recordcount does not work with a adOpenDynamic cursor
    rsLogDetails.Open sSql, conNewMACRO, adOpenKeyset, adLockReadOnly, adCmdText
    
    nLogNumber = rsLogDetails.RecordCount
    rsLogDetails.Close
    Set rsLogDetails = Nothing
    
    'Changed Mo Morris 6/4/01, truncate large messages that might have ben created by an error
    'to 255 characters so that it fits into field LogMessage in table LogDetails
    If Len(sMessage) > 255 Then
        sMessage = Left(sMessage, 255)
    End If
    
    'changed Mo Morris 6/1/00
    'check log message for single quotes
    sMessage = ReplaceQuotes(sMessage)
    
    'REM 31/10/02 - added timezone offset and Location
    nTimeZone = oTimezone.TimezoneOffset
    'Location will always be Local, will only be chnaged when transfered back to the server if a site
    sLocation = "Local"
    
    ' NCJ 4/2/00 - Changed dNow to sSQLNow
    'Mo Morris 30/8/01 Db Audit (UserId to UserName)
    'REM 31/10/02 - added TimeZone and Location
    sSql = " INSERT INTO LogDetails " _
        & "(LogDateTime,LogNumber,TaskId,LogMessage,UserName,LogDateTime_TZ,Location,Status)" _
        & " Values (" & sSQLNow & "," & nLogNumber & ",'" & sTaskId _
        & "','" & sMessage & "','" & sUserName & "'," & nTimeZone & ",'" & sLocation & "'," & 0 & ")"

    conNewMACRO.Execute sSql

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modSysDataXfer.gLogForXfer"
End Sub

'---------------------------------------------------------------------
Public Function GetNextPduMessageSql(ByRef conMACRO As ADODB.Connection, ByVal sSiteCode As String, _
                                ByVal sPduSaveDirectory As String, ByVal sLastMessageId As String, _
                                ByRef vErrorMsg As Variant) As String
'---------------------------------------------------------------------
' Mark last PDU message as being dealt with
' Read next PDU message
' write next PDU message to the file system
' return string formatted for asp
'---------------------------------------------------------------------
Dim sDblDate As String
Dim sSql As String
Dim rsPduMessage As ADODB.Recordset
Dim lMessageId As Long
Dim nMessageType As Integer
Dim sPduFileId As String
Dim sMessageBody As String
Dim sPduFileName As String
Dim sMessageToReturn As String

    On Error GoTo ErrHandler

    ' reset message to return
    sMessageToReturn = ""
    
    ' deal with previous message firstly
    If sLastMessageId > "" Then
        sDblDate = CStr(CDbl(Now))
        sDblDate = LocalNumToStandard(sDblDate)
        
        sSql = "UPDATE Message SET MessageReceived = 1, MessageReceivedTimeStamp = " & sDblDate
        sSql = sSql & " WHERE TrialSite = '" & sSiteCode & "' AND MessageId = " & sLastMessageId
        Call conMACRO.Execute(sSql)
    End If
    
    ' now get next message
    Set rsPduMessage = New ADODB.Recordset
    ' set sql
    sSql = "SELECT * FROM Message WHERE TrialSite = '" & sSiteCode & "' AND MessageReceived = 0 AND MessageDirection = 0  AND MessageType IN (50, 51)"
    ' open recordset
    rsPduMessage.Open sSql, conMACRO, adOpenForwardOnly, , adCmdText
    
    If rsPduMessage.EOF Then
        ' no more messages to receive
        lMessageId = -1
    Else
        ' get next message details
        ' messageid
        lMessageId = rsPduMessage("MessageId")
        ' message type
        nMessageType = rsPduMessage("MessageType")
        ' message parameters
        sPduFileId = rsPduMessage("MessageParameters")
        ' message body
        sMessageBody = rsPduMessage("MessageBody")
        ' only want 1st line of messagebody
        If InStr(sMessageBody, vbCrLf) > 0 Then
            sMessageBody = Left(sMessageBody, InStr(sMessageBody, vbCrLf))
        End If
    End If
    rsPduMessage.Close
    Set rsPduMessage = Nothing
    
    ' if next message retrieved write to file system
    If lMessageId > -1 Then
        ' firstly check if file already exists
        Select Case (nMessageType)
            Case 50:
                ' instruction
                sPduFileName = sPduFileId & ".pdi"
            Case 51:
                ' package
                sPduFileName = sPduFileId & ".pdu"
        End Select
        'check if file exists - only write if does not
        If Not FileExists(sPduSaveDirectory & sPduFileName) Then
            sPduFileName = SavePduFileToFileSystem(conMACRO, sPduFileId, sPduSaveDirectory, vErrorMsg)
        End If
        ' create message to return to asp
        ' messageid<br>filename
        sMessageToReturn = lMessageId & "<br>" & sPduFileName & "<br>" & sMessageBody
    Else
        sMessageToReturn = "."
    End If
    
    ' return message
    GetNextPduMessageSql = sMessageToReturn
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modSysDataXfer.GetNextPduMessageSql"
End Function

'---------------------------------------------------------------------
Public Function SavePduFileToFileSystem(ByRef conMACRO As ADODB.Connection, _
                                ByVal sPduFileId As String, _
                                ByVal sPduSaveDirectory As String, _
                                ByRef vErrorMsg As Variant) As String
'---------------------------------------------------------------------
' save file to PDU file system location
'---------------------------------------------------------------------
    Dim byteChunk() As Byte
    Dim sFullFilePath As String
    Dim sFileName As String
    Dim oFSO As New FileSystemObject
    Dim nFileNo As Integer
    ' filesize, noofchunks, leftoversize
    Dim lFileSize As Long
    Dim nNoOfChunks As Integer
    Dim nLeftOverSize As Integer
    Dim nChunkNo As Integer
    Const nCHUNKSIZE = 32768

    On Error GoTo ErrHandler
    
    ' get recordset
    Dim rsPduFile As ADODB.Recordset
    Set rsPduFile = New ADODB.Recordset
        
    ' set recordset object with sql to extract given fileid
    Dim sSql As String
    sSql = "SELECT * FROM PDUFILES WHERE FILEID = '" & sPduFileId & "'"
    ' reset filename
    sFileName = ""
    
    ' open recordset
    rsPduFile.Open sSql, conMACRO, adOpenForwardOnly, adLockReadOnly
    ' should just be the one record
    If Not rsPduFile.EOF Then
        ' set up file to write to
        ' fileid + extension
        sFileName = rsPduFile("FILEID").Value
        Select Case (CInt(rsPduFile("FILETYPE").Value))
            Case 0:
                sFileName = sFileName & ".pdi"
            Case 1:
                sFileName = sFileName & ".pdu"
        End Select
        sFullFilePath = sPduSaveDirectory & sFileName
     
        'check if file exists - only write if does not
        If Not FileExists(sFullFilePath) Then
            ' open file to write to
            nFileNo = FreeFile
            ' open file for write
            Open sFullFilePath For Binary Access Write As nFileNo
            
            'get filesize
            lFileSize = rsPduFile("FILEBINARY").ActualSize
            If lFileSize > 0 Then
                ' read in files in chunks
                nNoOfChunks = lFileSize \ nCHUNKSIZE
                ' fragment size
                nLeftOverSize = lFileSize Mod nCHUNKSIZE
                ReDim byteChunk(nLeftOverSize)
                byteChunk() = rsPduFile("FILEBINARY").GetChunk(nLeftOverSize)
                ' write 1st chunk to file
                Put nFileNo, , byteChunk()
                'loop through remaining chunks
                For nChunkNo = 1 To nNoOfChunks
                    ' redim buffer size
                    ReDim byteChunk(nCHUNKSIZE)
                    ' get chunk
                    byteChunk() = rsPduFile("FILEBINARY").GetChunk(nCHUNKSIZE)
                    ' write this chunk to file
                    Put nFileNo, , byteChunk()
                Next
            End If
            ' close file
            Close nFileNo
        End If
        ' close recordset
        rsPduFile.Close
        ' destroy
        Set rsPduFile = Nothing
    End If
    
    ' return filename
    SavePduFileToFileSystem = sFileName
    
    Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modSysDataXfer.SavePduFileToFileSystem"
End Function
