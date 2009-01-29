VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SysDataXfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999-2000. All Rights Reserved
'   File:       SysDataTransfer.cls
'   Author      Richard Meinesz, 11/11/02
'   Purpose:    Public routines to check, get and write system messages to the appropriate database and tables
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   DPH 23/01/2003 Added Getting / Saving report files
'   REM 23/03/04 - added LocalNumToStandard to GetSystemMessages
'   TA  18/01/2006 - MessageId now calculated by a sequence to avoid duplicate id problem
'----------------------------------------------------------------------------------------'
Option Explicit

Private mconMACRO As ADODB.Connection

'----------------------------------------------------------------------------------------'
Public Function MACRODBConnectionString(ByVal sDatabaseCode As String) As String
'----------------------------------------------------------------------------------------'
'REM 26/11/02
'Returns the MACRO database connection string
'----------------------------------------------------------------------------------------'
Dim sMessage As String

    MACRODBConnectionString = ConnectionString(sDatabaseCode, sMessage)

End Function

'----------------------------------------------------------------------------------------'
Public Function CheckSystemMessages(ByVal sUserName As String, ByVal sSiteCode As String, ByVal sDatabaseCode As String, _
                                    ByRef vErrorMsg As Variant, Optional sSecurityCon As String = "") As Boolean
'----------------------------------------------------------------------------------------'
'REM 11/11/02
'Checks to see if there are any system messages that have not been sent
'-------------------------------------------------------------------------
Dim conMACRO As ADODB.Connection
Dim vMessages As Variant
Dim oSecurityCon As ADODB.Connection

    On Error GoTo ErrLabel
    
    Set oSecurityCon = CreateSecurityConnection(CStr(vErrorMsg), sSecurityCon)
    
    'creats a routine level connecton to the MACRO database
    Set conMACRO = CreateConnection(oSecurityCon, sDatabaseCode, CStr(vErrorMsg))
    
    If CStr(vErrorMsg) <> "" Then GoTo ErrLabel
    
    'gets the user log and Log details messages and writes them to the message table, then will be retrieved as part of the system messages
    'will only do this if database is a site database
    Call GetUserLogMessages(oSecurityCon, conMACRO, sUserName)
    Call GetLogDetailsMessages(conMACRO, sUserName)
    
    vMessages = SystemMessages(conMACRO, sSiteCode)
    
    CheckSystemMessages = Not IsNull(vMessages)
    
    Set oSecurityCon = Nothing
    
Exit Function
ErrLabel:
    CheckSystemMessages = False
    If vErrorMsg = "" Then
        vErrorMsg = "Error while checking for system messages. " & "Error Description: " & Err.Description & ", Error Number: " & Err.Number
    End If
End Function

'----------------------------------------------------------------------------------------'
Public Function GetSystemMessages(ByVal sSiteCode As String, ByVal sDatabaseCode As String, _
                                  ByRef vErrorMsg As Variant, Optional sSecurityCon As String = "", _
                                  Optional ByVal vConfirmation As Variant = "") As String
'----------------------------------------------------------------------------------------'
'REM 11/11/02
'Gets all the system messages from the Message table that have a not yet been sent
'returns each message as a delimited string separated by "<p>"
' REM 23/03/04 - Added LocalNumToStandard around dates retieved from database
'----------------------------------------------------------------------------------------'
Dim oSecurityCon As ADODB.Connection
Dim conMACRO As ADODB.Connection
Dim sSql As String
Dim sMessageTypes As String
Dim rsMessage As ADODB.Recordset
Dim sSystemMessages As String
Dim vMessages As Variant
Dim sEncodedSysMessages As String
Dim vRestoreMessages As Variant
Dim i As Integer
Dim oCheckSum As IMEDCheckSum10.CheckSum
Dim sCheckSum As String
Dim nCount As Integer

    On Error GoTo ErrLabel
    
    'create connection to security database
    Set oSecurityCon = CreateSecurityConnection(CStr(vErrorMsg), sSecurityCon)
    If vErrorMsg <> "" Then GoTo ErrLabel
    'creats a routine level connecton to the MACRO database
    Set conMACRO = CreateConnection(oSecurityCon, sDatabaseCode, CStr(vErrorMsg))
    If vErrorMsg <> "" Then GoTo ErrLabel
    'confirm any previously sent messages
    If CStr(vConfirmation) <> "" Then
        Call ConfirmMessages(conMACRO, CStr(vConfirmation), CStr(vErrorMsg))
        If vErrorMsg <> "" Then GoTo ErrLabel
    End If
    
    'gets all the restore UserRole system messages that have not been sent yet
    vRestoreMessages = RestoreUserRoles(conMACRO, sSiteCode)
    
    'get all system messages from the message table that have not been sent yet
    vMessages = SystemMessages(conMACRO, sSiteCode)
    
    'if no messages then return ".", else get all the messages
    If IsNull(vMessages) And IsNull(vRestoreMessages) Then
        GetSystemMessages = gsEND_OF_MESSAGES
        
    Else
        sSystemMessages = ""
         'loop through the restore message types adding them to a delimited string
        If Not IsNull(vRestoreMessages) Then
            For i = 0 To UBound(vRestoreMessages, 2)
                ' REM 23/03/04 - Added LocalNumToStandard around dates retieved from database
                sSystemMessages = sSystemMessages & vRestoreMessages(eSDTMsgFields.fMessageId, i) _
                               & gsFIELDSEPARATOR & vRestoreMessages(eSDTMsgFields.fTrialSite, i) _
                               & gsFIELDSEPARATOR & vRestoreMessages(eSDTMsgFields.fClinicalTrialId, i) _
                               & gsFIELDSEPARATOR & vRestoreMessages(eSDTMsgFields.fMessageType, i) _
                               & gsFIELDSEPARATOR & vRestoreMessages(eSDTMsgFields.fUserName, i) _
                               & gsFIELDSEPARATOR & LocalNumToStandard(vRestoreMessages(eSDTMsgFields.fMessageTimeStamp, i)) _
                               & gsFIELDSEPARATOR & vRestoreMessages(eSDTMsgFields.fMessageTimeStamp_TZ, i) _
                               & gsFIELDSEPARATOR & vRestoreMessages(eSDTMsgFields.fMessageBody, i) _
                               & gsFIELDSEPARATOR & vRestoreMessages(eSDTMsgFields.fMessageParameters, i) _
                               & gsFIELDSEPARATOR & vRestoreMessages(eSDTMsgFields.fMessageDirection, i)
                'add message separator if not last message
                If i <> UBound(vRestoreMessages, 2) Then
                    sSystemMessages = sSystemMessages & gsMSGSEPARATOR
                End If
            Next
        End If
        
        ' then loop through the rest of the message types, adding them
        If Not IsNull(vMessages) Then
            'add message separator if the string is not empty
            If sSystemMessages <> "" Then
                sSystemMessages = sSystemMessages & gsMSGSEPARATOR
            End If
            
            For i = 0 To UBound(vMessages, 2)
                ' REM 23/03/04 - Added LocalNumToStandard around dates retieved from database
                sSystemMessages = sSystemMessages & vMessages(eSDTMsgFields.fMessageId, i) _
                               & gsFIELDSEPARATOR & vMessages(eSDTMsgFields.fTrialSite, i) _
                               & gsFIELDSEPARATOR & vMessages(eSDTMsgFields.fClinicalTrialId, i) _
                               & gsFIELDSEPARATOR & vMessages(eSDTMsgFields.fMessageType, i) _
                               & gsFIELDSEPARATOR & vMessages(eSDTMsgFields.fUserName, i) _
                               & gsFIELDSEPARATOR & LocalNumToStandard(vMessages(eSDTMsgFields.fMessageTimeStamp, i)) _
                               & gsFIELDSEPARATOR & vMessages(eSDTMsgFields.fMessageTimeStamp_TZ, i) _
                               & gsFIELDSEPARATOR & vMessages(eSDTMsgFields.fMessageBody, i) _
                               & gsFIELDSEPARATOR & vMessages(eSDTMsgFields.fMessageParameters, i) _
                               & gsFIELDSEPARATOR & vMessages(eSDTMsgFields.fMessageDirection, i)
                               
                'add message separator if not last message or string has exceeded max size
                If (i <> UBound(vMessages, 2)) And (Len(sSystemMessages) < 80000) Then
                    sSystemMessages = sSystemMessages & gsMSGSEPARATOR
                Else 'exit the loop as end of messages or reached max of 80 000 chars
                    Exit For
                End If
            Next
            
        End If
        
        'do checksum on final string
        Set oCheckSum = New CheckSum
        sCheckSum = oCheckSum.GetStringCheckSum(sSystemMessages)
        'add checksum to string
        sSystemMessages = sSystemMessages & gsCHECKSUM_SEPARATOR & sCheckSum
        
        'Replace any invalid characters in the string
        sEncodedSysMessages = HexEncodeChars(sSystemMessages, gsCHARS_TO_ENCODE)
    
        GetSystemMessages = sEncodedSysMessages
        
        Set oCheckSum = Nothing
        
    End If
    
    Set oSecurityCon = Nothing
    Set conMACRO = Nothing
    
Exit Function
ErrLabel:
    If vErrorMsg = "" Then
        vErrorMsg = "Error while getting system messages. Error Description: " & Err.Description & ", Error Number: " & Err.Number
    End If
End Function

'----------------------------------------------------------------------------------------'
Public Function WriteSystemMessage(ByVal sDatabaseCode As String, ByVal sEncodedSysMessages As String, _
                                   ByRef vErrorMsg As Variant, Optional sSecurityCon As String = "") As String
'----------------------------------------------------------------------------------------'
'REM 11/11/02
'Writes a system message to the Message tabel and writes the data to the correct datatabase and table
'and returns the message's Id as confirmation of success
'----------------------------------------------------------------------------------------'
Dim oSecurityCon As ADODB.Connection
Dim conMACRO As ADODB.Connection
Dim vMessage As Variant
Dim lMessageId As Long
Dim sTrialSite As String
Dim lClinicalTrialId As Long
Dim nMessageType As Integer
Dim sMsgUserName As String
Dim sMessageTimeStamp As String
Dim nMessageTimepStamp_TZ As Integer
Dim sMessageBody As String
Dim sMessageParameters As String
Dim nMessageDirection As Integer
Dim bWrite As Boolean
Dim sNextMessage As String
Dim vSystemMessages As Variant
Dim sMessageIds As String
Dim sReturn As String
Dim sSystemMessages As String
Dim i As Integer
Dim nMessageReceived As Integer
Dim vCheckSum As Variant
Dim sCheckSum As String
Dim oCheckSum As CheckSum
Dim sErrorMsg As String

    On Error GoTo ErrLabel

    sSystemMessages = HexDecodeChars(sEncodedSysMessages)
    
    'split message to get the checksum value and message
    vCheckSum = Split(sSystemMessages, gsCHECKSUM_SEPARATOR)
    sSystemMessages = vCheckSum(0)
    sCheckSum = vCheckSum(1)
    
    'do check on the checksum to see if data has been corrupted
    Set oCheckSum = New CheckSum
    
    'if checksum fails then return error message ans exit function
    If Not oCheckSum.CheckStringCheckSum(sSystemMessages, sCheckSum) Then
        vErrorMsg = gsCORRUPT_DATA
        WriteSystemMessage = ""
        Exit Function
    End If

    'Pick up if an error has occurred in the ASP page
    If Left(sSystemMessages, 5) = "ERROR" Or InStr(1, LCase(sSystemMessages), "<html") > 0 Then
        vErrorMsg = "System message could not be downloaded"
        WriteSystemMessage = ""
        Exit Function
    End If
    
    WriteSystemMessage = ""
    
    'create connection to secuirty database
    Set oSecurityCon = CreateSecurityConnection(sErrorMsg, sSecurityCon)
    If sErrorMsg <> "" Then GoTo ErrLabel
    
    'creats a routine level connecton to the MACRO database
    Set conMACRO = CreateConnection(oSecurityCon, sDatabaseCode, sErrorMsg)
    If sErrorMsg <> "" Then GoTo ErrLabel
    
    'get each system message
    vSystemMessages = Split(sSystemMessages, gsMSGSEPARATOR)
    
    'loop through all the messages, writing them to appropriate database and tabel
    For i = 0 To UBound(vSystemMessages)
    
        'get next message
        sNextMessage = vSystemMessages(i)

        'get each parameter
        vMessage = Split(sNextMessage, gsFIELDSEPARATOR)
        
        lMessageId = vMessage(eSDTMessage.mMessageId)
        sTrialSite = vMessage(eSDTMessage.mTrialSite)
        lClinicalTrialId = vMessage(eSDTMessage.mClinicalTrialId)
        nMessageType = vMessage(eSDTMessage.mMessageType)
        sMsgUserName = vMessage(eSDTMessage.mUserName)
        sMessageTimeStamp = vMessage(eSDTMessage.mMessageTimeStamp)
        nMessageTimepStamp_TZ = vMessage(eSDTMessage.mMessageTimeStamp_TZ)
        sMessageBody = vMessage(eSDTMessage.mMessageBody)
        sMessageParameters = vMessage(eSDTMessage.mMessageParameters)
        nMessageDirection = vMessage(eSDTMessage.mMessageDirection)

        'select on message type and write away message to correct database and table
        Select Case nMessageType
        
        Case ExchangeMessageType.User

            bWrite = WriteUserDetails(conMACRO, oSecurityCon, sSecurityCon, sTrialSite, nMessageType, sMsgUserName, sMessageParameters, sMessageBody, nMessageReceived, sErrorMsg)
        
        Case ExchangeMessageType.UserRole, ExchangeMessageType.RestoreUserRole

            bWrite = WriteUserRole(conMACRO, oSecurityCon, sTrialSite, sDatabaseCode, nMessageType, sMsgUserName, sMessageParameters, sMessageBody, nMessageReceived, sErrorMsg)
        
        Case ExchangeMessageType.PasswordChange

            bWrite = WriteChangePassword(conMACRO, oSecurityCon, sTrialSite, nMessageType, sMsgUserName, sMessageParameters, sMessageBody, nMessageReceived, sErrorMsg)
        
        Case ExchangeMessageType.Role

            bWrite = WriteRole(oSecurityCon, sMessageParameters, sErrorMsg)
            nMessageReceived = -CInt(bWrite)
        
        Case ExchangeMessageType.SystemLog

            bWrite = WriteSystemLog(conMACRO, sMessageParameters, sTrialSite, sErrorMsg)
            nMessageReceived = -CInt(bWrite)
            
        Case ExchangeMessageType.UserLog

            bWrite = WriteUserLog(oSecurityCon, sMessageParameters, sTrialSite, sErrorMsg)
            nMessageReceived = -CInt(bWrite)
            
        Case ExchangeMessageType.PasswordPolicy
            
            bWrite = WritePswdPolicy(oSecurityCon, sMessageParameters, sErrorMsg)
            nMessageReceived = -CInt(bWrite)
            
        End Select

        'if data was successfully written to the database then add the message to the message table
        'and return the message id, message received time stamp and time-zone offset as confirmation of message sent
        If bWrite Then
            sReturn = WriteMsgToMessageTable(conMACRO, sTrialSite, lClinicalTrialId, nMessageType, sMessageTimeStamp, sMsgUserName, sMessageBody, _
                                        sMessageParameters, nMessageDirection, lMessageId, nMessageTimepStamp_TZ, nMessageReceived, sErrorMsg)
            
            If sReturn = "" Then GoTo ErrLabel
            
            WriteSystemMessage = WriteSystemMessage & sReturn
            
            If i <> UBound(vSystemMessages) Then
                WriteSystemMessage = WriteSystemMessage & gsSEPARATOR
            End If
            
        Else ' return "" to indcate failure to write data
            vErrorMsg = sErrorMsg & vbCrLf & "Error while writing message: MessageId = " & lMessageId & ", MessageBody: " & sMessageBody & vbCrLf & vbCrLf
            WriteSystemMessage = ""
        End If
    Next
    
    Set oSecurityCon = Nothing
    Set conMACRO = Nothing

Exit Function
ErrLabel:
    If sErrorMsg = "" Then
        vErrorMsg = "Error while writing system messages to the database.  Error description: " & Err.Description & ", Error Number: " & Err.Number
    Else
        vErrorMsg = sErrorMsg
    End If
    WriteSystemMessage = ""
End Function

'----------------------------------------------------------------------------------------'
Public Function GetNewPassword(ByVal sUserName As String, ByVal sPassword As String, ByVal sDatabaseCode As String, _
                                  ByVal sSiteCode As String, ByRef vErrorMsg As Variant, _
                                  Optional ByVal vConfirmation As Variant = "") As String
'----------------------------------------------------------------------------------------'
'REM 11/11/02
'Gets the new password and other user details from the server database,
'used when a site user has forgotten their password and it is reset on the server database
'----------------------------------------------------------------------------------------'
Dim oSecurityCon As ADODB.Connection
Dim conMACRO As ADODB.Connection
Dim sSql As String
Dim rsUser As ADODB.Recordset
Dim sHashedServerPassword As String
Dim sHashedUserPassword As String
Dim sPasswordAndDate As String
Dim sPasswordCreated As String
Dim sChangePassword As String
Dim i As Integer
Dim oCheckSum As IMEDCheckSum10.CheckSum
Dim sCheckSum As String
Dim sEncodedPswdMessages As String
Dim vPassword As Variant
Dim rsPswd As ADODB.Recordset
Dim nPswdRetries As Integer
Dim bSysAdmin As Boolean
Dim nFailedAttempts As Integer
Dim oUser As MACROUser
Dim sSecCon As String

    On Error GoTo ErrLabel
    sSecCon = ""
    'create connection to secuirty database
    Set oSecurityCon = CreateSecurityConnection(CStr(vErrorMsg), sSecCon)
    If vErrorMsg <> "" Then GoTo ErrLabel
    'creats a routine level connecton to the MACRO database
    Set conMACRO = CreateConnection(oSecurityCon, sDatabaseCode, CStr(vErrorMsg))
    If vErrorMsg <> "" Then GoTo ErrLabel
    
    'confirm previously sent password message and exit
    If CStr(vConfirmation) <> "" Then
        Call ConfirmMessages(conMACRO, CStr(vConfirmation), CStr(vErrorMsg))
        GetNewPassword = gsEND_OF_MESSAGES
        Exit Function
    End If
    
    'get User details
'    sSQL = "SELECT * FROM MACROUser" _
'        & " WHERE UserName = '" & sUserName & "'"
'        Set rsUser = New ADODB.Recordset
'        rsUser.Open sSQL, oSecurityCon, adOpenKeyset, adLockReadOnly, adCmdText
        
    'REM 21/01/03 - check user name in Uppercase in Oracle as it is case sensitive
    Select Case Connection_Property(CONNECTION_PROVIDER, sSecCon)
    Case CONNECTION_MSDAORA, CONNECTION_ORAOLEDB_ORACLE
        sSql = "SELECT * FROM MACROUser WHERE upper(UserName) = upper('" & sUserName & "')"
    Case Else
        sSql = "SELECT * FROM MACROUser WHERE UserName = '" & sUserName & "'"
    End Select
    Set rsUser = New ADODB.Recordset
    rsUser.Open sSql, oSecurityCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    'check how many times a user can attempt to login (if 0 then this propery is not set)
    sSql = "SELECT PASSWORDRETRIES FROM MACROPASSWORD"
    Set rsPswd = New ADODB.Recordset
    rsPswd.Open sSql, oSecurityCon, adOpenKeyset, adLockReadOnly, adCmdText
    nPswdRetries = rsPswd!PasswordRetries


    If rsUser.RecordCount <> 0 Then
        'see if user is sys admin
        bSysAdmin = (rsUser!SysAdmin = 1)
        
        'get number of failed attempts
        nFailedAttempts = rsUser!FailedAttempts
        'If PswdRetries  > 0  then property is set, so check if failed attempts is > retries, and see if user is sysadmin, if so can't be locked out
        If (nPswdRetries <> 0) And (nFailedAttempts >= nPswdRetries) And (bSysAdmin = False) Then
            'account has been locked out
            GetNewPassword = "Account Locked out"
            vErrorMsg = "User account has been locked on the server.  User attempted to login with incorrect password more than " & nPswdRetries & " times."
            
        Else 'check user details
        
            sHashedServerPassword = rsUser!UserPassword
            'REM 24/03/04 - ensure string contains dots not commas
            sPasswordCreated = LocalNumToStandard(rsUser!PasswordCreated)
            
            sPasswordAndDate = sPassword & sPasswordCreated
            
            sHashedUserPassword = HashHexEncodeString(sPasswordAndDate)
            
            'check if passwords match
            If sHashedServerPassword = sHashedUserPassword Then
                'get all system messages from the message table that are for change password and user details
                'that have not been sent yet
                vPassword = UserPasswordAndDetails(conMACRO, sSiteCode)
                
                'if no messages then return ".", else get all the messages
                If IsNull(vPassword) Then
                    GetNewPassword = gsEND_OF_MESSAGES
                    vErrorMsg = "There are no new user details to download"
                Else
                    'loop through all user details
                    For i = 0 To UBound(vPassword, 2)
                        sChangePassword = sChangePassword & vPassword(eSDTMsgFields.fMessageId, i) _
                                       & gsFIELDSEPARATOR & vPassword(eSDTMsgFields.fTrialSite, i) _
                                       & gsFIELDSEPARATOR & vPassword(eSDTMsgFields.fClinicalTrialId, i) _
                                       & gsFIELDSEPARATOR & vPassword(eSDTMsgFields.fMessageType, i) _
                                       & gsFIELDSEPARATOR & vPassword(eSDTMsgFields.fUserName, i) _
                                       & gsFIELDSEPARATOR & vPassword(eSDTMsgFields.fMessageTimeStamp, i) _
                                       & gsFIELDSEPARATOR & vPassword(eSDTMsgFields.fMessageTimeStamp_TZ, i) _
                                       & gsFIELDSEPARATOR & vPassword(eSDTMsgFields.fMessageBody, i) _
                                       & gsFIELDSEPARATOR & vPassword(eSDTMsgFields.fMessageParameters, i) _
                                       & gsFIELDSEPARATOR & vPassword(eSDTMsgFields.fMessageDirection, i)
                    
                        If i <> UBound(vPassword, 2) Then
                            sChangePassword = sChangePassword & gsMSGSEPARATOR
                        End If
                    Next
                    
                    'create a checksum of the message
                    Set oCheckSum = New CheckSum
                    sCheckSum = oCheckSum.GetStringCheckSum(sChangePassword)
                    
                    'add the checksum to the message
                    sChangePassword = sChangePassword & gsCHECKSUM_SEPARATOR & sCheckSum
                    
                    'Replace any invalid characters in the string
                    sEncodedPswdMessages = HexEncodeChars(sChangePassword, gsCHARS_TO_ENCODE)
                
                    GetNewPassword = sEncodedPswdMessages
                    
                    Set oCheckSum = Nothing
                    
                End If
            
            Else 'password does not match
            
                'check if user is a sys admin, if not can be locked out
                If Not bSysAdmin Then
                    
                    'does require account lockout
                    If nPswdRetries > 0 Then
                        
                        'add one
                        nFailedAttempts = nFailedAttempts + 1
                        'update database
                        sSql = "UPDATE MACROUser SET FailedAttempts = " & nFailedAttempts _
                            & " WHERE UserName = '" & sUserName & "'"
                        oSecurityCon.Execute sSql, , adCmdText
                        
                        'check to see if user has exceeded max number of retries
                        If nFailedAttempts >= nPswdRetries Then
                            GetNewPassword = "Account Locked out"
                            vErrorMsg = "User account has been locked on the server.  User attempted to login with incorrect password more than " & nPswdRetries & " times."
                            Set oUser = New MACROUser
                            Call oUser.gLog(sUserName, "Login", "Forgotten password login from remote site failed, user locked out", oSecurityCon)
                            Set oUser = Nothing
                        Else
                            GetNewPassword = ""
                            vErrorMsg = "Incorrect password"
                        End If
                        
                    Else
                        GetNewPassword = ""
                        vErrorMsg = "Incorrect password"
                    End If
    
                Else
                    GetNewPassword = ""
                    vErrorMsg = "Incorrect password"
                End If
            End If
        End If
        
    Else
            GetNewPassword = ""
            vErrorMsg = "User name does not exist"
    End If
    
    rsUser.Close
    Set rsUser = Nothing
    
    rsPswd.Close
    Set rsPswd = Nothing
    
    Set conMACRO = Nothing
    Set oSecurityCon = Nothing

Exit Function
ErrLabel:
    GetNewPassword = ""
    vErrorMsg = "Error in retrieving new user password. Error description: " & Err.Description & ", Error Number: " & Err.Number
End Function

'----------------------------------------------------------------------------------------'
Public Function WriteNewPassword(ByVal sDatabaseCode As String, ByVal sEncodedSysMessages As String, _
                                 ByRef vErrorMsg As Variant) As String
'----------------------------------------------------------------------------------------'
'REM 06/12/02
'Writes the new password and other user info to the security database
'Used when a Site user has forgotten their password and the new one is retrieved by the GetNewPassword function
'and written here.
'----------------------------------------------------------------------------------------'
Dim bChangedPassword As Boolean
Dim oSecurityCon As ADODB.Connection
Dim conMACRO As ADODB.Connection
Dim vCheckSum As Variant
Dim sCheckSum As String
Dim oCheckSum As CheckSum
Dim i As Integer
Dim nMessageReceived As Integer
Dim lMessageId As Long
Dim sTrialSite As String
Dim lClinicalTrialId As Long
Dim nMessageType As Integer
Dim sMsgUserName As String
Dim sMessageTimeStamp As String
Dim nMessageTimepStamp_TZ As Integer
Dim sMessageBody As String
Dim sMessageParameters As String
Dim nMessageDirection As Integer
Dim sSystemMessages As String
Dim vSystemMessages As Variant
Dim sNextMessage As String
Dim vMessage As Variant
Dim sReturn As String
Dim sSecCon As String

    On Error GoTo ErrLabel

    sSystemMessages = HexDecodeChars(sEncodedSysMessages)
    
    'split message to get the checksum value and message
    vCheckSum = Split(sSystemMessages, gsCHECKSUM_SEPARATOR)
    sSystemMessages = vCheckSum(0)
    sCheckSum = vCheckSum(1)
    
    'do check on the checksum to see if data has been corrupted
    Set oCheckSum = New CheckSum
    
    'if checksum fails then return error message and exit function
    If Not oCheckSum.CheckStringCheckSum(sSystemMessages, sCheckSum) Then
        vErrorMsg = gsCORRUPT_DATA
        WriteNewPassword = ""
        Exit Function
    End If

    'create connection to secuirty database
    Set oSecurityCon = CreateSecurityConnection(CStr(vErrorMsg), sSecCon)
    If vErrorMsg <> "" Then GoTo ErrLabel
    
    'creats a routine level connecton to the MACRO database
    Set conMACRO = CreateConnection(oSecurityCon, sDatabaseCode, CStr(vErrorMsg))
    If vErrorMsg <> "" Then GoTo ErrLabel
    
    'get each system message
    vSystemMessages = Split(sSystemMessages, gsMSGSEPARATOR)
    
    'loop through all the messages, writing them to appropriate database and tabel
    For i = 0 To UBound(vSystemMessages)
    
        'get next message
        sNextMessage = vSystemMessages(i)

        'get each parameter
        vMessage = Split(sNextMessage, gsFIELDSEPARATOR)
        
        lMessageId = vMessage(eSDTMessage.mMessageId)
        sTrialSite = vMessage(eSDTMessage.mTrialSite)
        lClinicalTrialId = vMessage(eSDTMessage.mClinicalTrialId)
        nMessageType = vMessage(eSDTMessage.mMessageType)
        sMsgUserName = vMessage(eSDTMessage.mUserName)
        sMessageTimeStamp = vMessage(eSDTMessage.mMessageTimeStamp)
        nMessageTimepStamp_TZ = vMessage(eSDTMessage.mMessageTimeStamp_TZ)
        sMessageBody = vMessage(eSDTMessage.mMessageBody)
        sMessageParameters = vMessage(eSDTMessage.mMessageParameters)
        nMessageDirection = vMessage(eSDTMessage.mMessageDirection)

        'select on message type and write away message to correct database and table
        Select Case nMessageType

        Case ExchangeMessageType.PasswordChange
            bChangedPassword = WriteChangePassword(conMACRO, oSecurityCon, sTrialSite, nMessageType, sMsgUserName, sMessageParameters, sMessageBody, nMessageReceived, CStr(vErrorMsg))
        Case ExchangeMessageType.User
            bChangedPassword = WriteUserDetails(conMACRO, oSecurityCon, sSecCon, sTrialSite, nMessageType, sMsgUserName, sMessageParameters, sMessageBody, nMessageReceived, CStr(vErrorMsg))
        End Select
        
        'if data was successfully written to the database then add the message to the message table
        'and return the message id, message received time stamp and time-zone offset as confirmation of message sent
        If bChangedPassword Then
            sReturn = WriteMsgToMessageTable(conMACRO, sTrialSite, lClinicalTrialId, nMessageType, sMessageTimeStamp, sMsgUserName, sMessageBody, _
                                        sMessageParameters, nMessageDirection, lMessageId, nMessageTimepStamp_TZ, nMessageReceived, CStr(vErrorMsg))
            
            WriteNewPassword = WriteNewPassword & sReturn
            
            If i <> UBound(vSystemMessages) Then
                WriteNewPassword = WriteNewPassword & gsSEPARATOR
            End If
            
        Else ' return "" to indcate failure to write data
            WriteNewPassword = ""
            Exit For
        End If

    Next

    Set oSecurityCon = Nothing
    Set conMACRO = Nothing

Exit Function
ErrLabel:
    vErrorMsg = "Error while writing new password messages to the database.  Error description: " & Err.Description & ", Error Number: " & Err.Number
    WriteNewPassword = ""
End Function

'----------------------------------------------------------------------------------------'
Public Function WriteReportFiles(ByVal sDatabaseCode As String, ByVal sZipFileName As String, _
                                 ByRef vErrorMsg As Variant, ByVal sUserName As String, _
                                 ByVal sSiteCode As String, Optional sSecurityCon As String = "") As Boolean
'----------------------------------------------------------------------------------------'
' Store the Report files to the remote site's reports folder
' Unzip the ZIPPED file & copy files
'----------------------------------------------------------------------------------------'
Dim conMACRO As ADODB.Connection
Dim oSecurityCon As ADODB.Connection
Dim bUnzippedAndStored As Boolean

    On Error GoTo ErrLabel
    
    Set oSecurityCon = CreateSecurityConnection(CStr(vErrorMsg), sSecurityCon)
    
    'creates a routine level connection to the MACRO database
    Set conMACRO = CreateConnection(oSecurityCon, sDatabaseCode, CStr(vErrorMsg))
    
    If vErrorMsg <> "" Then GoTo ErrLabel

    ' write report files to relevant path after extracting from ZIP file
    bUnzippedAndStored = StoreReportFiles(conMACRO, oSecurityCon, sDatabaseCode, sSiteCode, sZipFileName)
    
    ' if no filename then error
    If Not bUnzippedAndStored Then
        vErrorMsg = "Error whilst writing report file updates. Failed to create report ZIP file. "
    End If
    
    ' Close connections
    If oSecurityCon.State = adStateOpen Then
        oSecurityCon.Close
    End If
    If conMACRO.State = adStateOpen Then
        conMACRO.Close
    End If
    
    Set oSecurityCon = Nothing
    Set conMACRO = Nothing
    
    WriteReportFiles = bUnzippedAndStored
    
    Exit Function
ErrLabel:
    WriteReportFiles = False
    If vErrorMsg = "" Then
        vErrorMsg = "Error whilst writing report file updates. " & "Error Description: " & Err.Description & ", Error Number: " & Err.Number
    End If
End Function

'----------------------------------------------------------------------------------------'
Public Function GetReportFiles(ByVal sSiteCode As String, ByVal sDatabaseCode As String, _
                                ByVal sUserName As String, sLastTransDate As String, _
                                ByRef vErrorMsg As Variant, Optional sSecurityCon As String = "") As String
'----------------------------------------------------------------------------------------'
' On the server create a zip file containing all modified report files
' return filename of zip file
'----------------------------------------------------------------------------------------'
'REM 13/10/03 - add sLastTransDate to routine
'----------------------------------------------------------------------------------------'
Dim conMACRO As ADODB.Connection
Dim oSecurityCon As ADODB.Connection
Dim sFileName As String
Dim bNoFiles As Boolean
Dim sMessageParameters As String
Dim sMACROConn As String

    On Error GoTo ErrLabel
    
    ' initialise vErrorMsg
    vErrorMsg = ""
    
    Set oSecurityCon = CreateSecurityConnection(CStr(vErrorMsg), sSecurityCon)
    
    'creates a routine level connection to the MACRO database
    Set conMACRO = CreateConnection(oSecurityCon, sDatabaseCode, CStr(vErrorMsg), sMACROConn)
    
    If vErrorMsg <> "" Then GoTo ErrLabel

    ' create filename
    bNoFiles = True
    sFileName = CreateReportFilesZIP(conMACRO, oSecurityCon, sUserName, sDatabaseCode, sSiteCode, sLastTransDate, bNoFiles, sMACROConn)
    
    ' if no filename then error
    If sFileName = "" Then
        If Not bNoFiles Then
            vErrorMsg = "Error while checking for report file updates. Failed to create report ZIP file. "
        Else
            Call gLogForXfer("ReportXferServer", "Site " & sSiteCode & " has no reports to download ", conMACRO, sUserName)
        End If
    Else
        ' write log entry to server
        ' put together required parameters
        Call gLogForXfer("ReportXferServer", "Site " & sSiteCode & " created report ZIP file " & sFileName, conMACRO, sUserName)
    End If
    
    ' set filename function return
    GetReportFiles = sFileName
    
    ' Close connections
    If oSecurityCon.State = adStateOpen Then
        oSecurityCon.Close
    End If
    If conMACRO.State = adStateOpen Then
        conMACRO.Close
    End If
    
    Set oSecurityCon = Nothing
    Set conMACRO = Nothing
    Exit Function
ErrLabel:
    GetReportFiles = ""
    If vErrorMsg = "" Then
        vErrorMsg = "Error while checking for report file updates. " & "Error Description: " & Err.Description & ", Error Number: " & Err.Number
    End If
End Function

'----------------------------------------------------------------------------------------'
Public Function ConfirmReportFiles(ByVal sSiteCode As String, ByVal sDatabaseCode As String, _
                                ByVal sUserName As String, ByRef vErrorMsg As Variant, Optional sSecurityCon As String = "") As String
'----------------------------------------------------------------------------------------'
' On the server confirm the receipt of the report files
'----------------------------------------------------------------------------------------'
Dim conMACRO As ADODB.Connection
Dim oSecurityCon As ADODB.Connection

    On Error GoTo ErrLabel
    
    ' initialise vErrorMsg
    vErrorMsg = ""
    
    Set oSecurityCon = CreateSecurityConnection(CStr(vErrorMsg), sSecurityCon)
    
    'creates a routine level connection to the MACRO database
    Set conMACRO = CreateConnection(oSecurityCon, sDatabaseCode, CStr(vErrorMsg))
    
    If vErrorMsg <> "" Then GoTo ErrLabel

    ' write log entry to server
    ' put together required parameters
    Call gLogForXfer("ReportXfer", "Site " & sSiteCode & " confirmed receipt of modified report files", conMACRO, sUserName)

    ' Close connections
    If oSecurityCon.State = adStateOpen Then
        oSecurityCon.Close
    End If
    If conMACRO.State = adStateOpen Then
        conMACRO.Close
    End If
    
    Set oSecurityCon = Nothing
    Set conMACRO = Nothing
    
    ConfirmReportFiles = "SUCCESS"
    
    Exit Function
ErrLabel:
    ConfirmReportFiles = "ERROR"
    If vErrorMsg = "" Then
        vErrorMsg = "Error while confirming report file updates. " & "Error Description: " & Err.Description & ", Error Number: " & Err.Number
    End If
End Function

'----------------------------------------------------------------------------------------'
Public Function GetNextPduMessage(ByVal sSiteCode As String, ByVal sDatabaseCode As String, _
                                ByVal sPduSaveDirectory As String, ByVal sLastMessage As Variant, _
                                ByRef vErrorMsg As Variant) As String
'----------------------------------------------------------------------------------------'
' Get next PDU message from database, storing it on the filesystem
'----------------------------------------------------------------------------------------'
Dim conMACRO As ADODB.Connection
Dim oSecurityCon As ADODB.Connection
Dim sConnection As String
Dim sErrorMsg As String

    On Error GoTo ErrLabel
    
    Set oSecurityCon = CreateSecurityConnection(CStr(vErrorMsg), "")
    
    ' check for errors
    If vErrorMsg <> "" Then GoTo ErrLabel
 
    ' get connectionstring to db
    sConnection = ConnectionString(sDatabaseCode, sErrorMsg, oSecurityCon)
    
    ' check for errors
    vErrorMsg = sErrorMsg
    If vErrorMsg <> "" Then GoTo ErrLabel
   
    ' if an oracle MSDAORA connection change to ORAOLEDB.ORACLE connection as require BLOBs support
    If InStr(1, Connection_Property(CONNECTION_PROVIDER, sConnection), CONNECTION_MSDAORA) > 0 Then
        ' reconstruct connection string as ORAOLEDB.ORACLE
        sConnection = Connection_String(CONNECTION_ORAOLEDB_ORACLE, Connection_Property(CONNECTION_DATASOURCE, sConnection), _
                "", Connection_Property(CONNECTION_USERID, sConnection), Connection_Property(CONNECTION_PASSWORD, sConnection))
    End If

    ' create connection object
    If sConnection <> "" Then
        'create connection for selected database
        Set conMACRO = New ADODB.Connection
    
        conMACRO.Open sConnection
        conMACRO.CursorLocation = adUseClient
    End If
    
    ' write/read to/from db and extract file from database
    GetNextPduMessage = GetNextPduMessageSql(conMACRO, sSiteCode, sPduSaveDirectory, sLastMessage, vErrorMsg)
    
    ' if no filename then error
    If GetNextPduMessage = "" Then
        vErrorMsg = "Error whilst retrieving and creating next PDU file. "
    End If
    
    ' Close connections
    If oSecurityCon.State = adStateOpen Then
        oSecurityCon.Close
    End If
    If conMACRO.State = adStateOpen Then
        conMACRO.Close
    End If
    
    Set oSecurityCon = Nothing
    Set conMACRO = Nothing
        
    Exit Function
ErrLabel:
    GetNextPduMessage = ""
    If vErrorMsg = "" Then
        vErrorMsg = "Error whilst retrieving and creating next PDU file. " & "Error Description: " & Err.Description & ", Error Number: " & Err.Number
    End If
End Function
'----------------------------------------------------------------------
Public Function GetNextMessageId(dbConnection As Connection) As Long
'----------------------------------------------------------------------
'   TA  18/01/2006 - MessageId now calculated by a sequence to avoid duplicate id problem
'----------------------------------------------------------------------
On Error GoTo ErrLabel

     GetNextMessageId = CLng(GetNextSequenceNo(dbConnection, "SEQ_MESSAGEID"))

Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "SysDataXfer.GetNextMessageId"
End Function


'----------------------------------------------------------------------
Private Function GetNextSequenceNo(dbConnection As Connection, sSequenceName As String) As Long
'----------------------------------------------------------------------
'Get next value in a sequence
' the most sql server can return is a long
' this should be moved to a libADO when refactored
'----------------------------------------------------------------------
Dim vData As Variant
Dim oCom As Command
Dim rs As Recordset
On Error GoTo ErrLabel
    
    If InStr(dbConnection.ConnectionString, CONNECTION_SQLOLEDB) = 0 Then
        'oracle
        Set rs = New Recordset
        rs.Open "select " & sSequenceName & ".nextval from dual", dbConnection
        vData = rs.GetRows
        rs.Close
        GetNextSequenceNo = vData(0, 0)
    Else
        'sql server
        Set oCom = New Command
        oCom.CommandText = "sp_MACRO_Seq_" & sSequenceName
        oCom.Parameters.Append oCom.CreateParameter("@NextVal", adInteger, adParamOutput)
        Set oCom.ActiveConnection = dbConnection
        oCom.CommandType = adCmdStoredProc
        oCom.Execute
        GetNextSequenceNo = oCom.Parameters("@NextVal")
        Set oCom = Nothing
    End If
    Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "SysDataXfer.GetNextSequenceNo"
End Function


