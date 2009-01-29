VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChangeUserDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------
' File: ChangeUserDetails.cls
' Copyright: InferMed Ltd 2004 All Rights Reserved
' Author: Toby Aldridge, September 2004
' Purpose: Store the properties of a suer to be changed by web client
'------------------------------------------------------------------
' REVISIONS

'------------------------------------------------------------------

Option Explicit

'from frmUserDetails
Private Enum eUserDetails
    udNewUser = 0
    udEditUser = 1
    udDisableUser = 2
End Enum

Private Const gsCHANGE_PSWD = "ChangePassword" 'from basCommon
Private Const gsUSER_ENABLED = "UserEnabled" 'from basCommon
Private Const gsUSER_DISABLED = "UserDisabled" 'from basCommon
Private Const gsUSER_UNLOCKED = "UserUnLocked"  'from basCommon
Private Const gsCHANGE_USERNAME_FULL = "ChangeUserNameFull" 'from basCommon

'instance of logger
Private moLog As APILogger

'-------------------------------------------------------------------------------------
Private Function GetCurrentUser(sSerialisedUser As String) As MACROUser
'-------------------------------------------------------------------------------------
'get user object from serialised string
'-------------------------------------------------------------------------------------
Dim oUser As MACROUser

    ' Create the MACRO User
    Set oUser = New MACROUser
    moLog.Log "creating user from serialised state"
    Call oUser.SetStateHex(sSerialisedUser)
    Set GetCurrentUser = oUser

End Function

'-------------------------------------------------------------------------------------
Private Sub Log(sMessage As String)
'-------------------------------------------------------------------------------------
'log to logger
'-------------------------------------------------------------------------------------

    moLog.Log sMessage
    
End Sub

'-------------------------------------------------------------------------------------
Public Function UpdateUser(sSerialisedUser As String, oNewDetail As UserDetail, ByRef vMessage As Variant) As Boolean
'-------------------------------------------------------------------------------------
'update user details
'-------------------------------------------------------------------------------------
Dim oCurrentUser As MACROUser
Dim conSecurity As Connection
Dim conMACRO As Connection
Dim sFailMessage As String

    sFailMessage = ""
    Set oCurrentUser = GetCurrentUser(sSerialisedUser)
    Log "creating security connection"
    Set conSecurity = New Connection
    Call conSecurity.Open(SecurityDatabasePath)
    conSecurity.CursorLocation = adUseClient
    Log "creating MACRO connection"
    Set conMACRO = New Connection
    conMACRO.Open (oCurrentUser.CurrentDBConString)
    conMACRO.CursorLocation = adUseClient
    
    UpdateUser = UpdateUserDetail(oNewDetail, oCurrentUser, sFailMessage, conMACRO, conSecurity)
    
    conSecurity.Close
    conMACRO.Close
    If sFailMessage = "" Then
        vMessage = moLog.Messages
    Else
        vMessage = sFailMessage
    End If
    
End Function

'----------------------------------------------------------------------------------------'
Private Property Get SecurityDatabasePath() As String
'----------------------------------------------------------------------------------------'
'return security db path
'-------------------------------------------------------------------------------------
    SecurityDatabasePath = GetMACROSetting("SecurityPath", "")
    If SecurityDatabasePath <> "" Then
        SecurityDatabasePath = DecryptString(SecurityDatabasePath)
    End If

End Property

'-------------------------------------------------------------------------------------
Public Function GetUsersDetails(sSerialisedUser As String, sUserName As String, ByRef sMessage As Variant) As Collection
'-------------------------------------------------------------------------------------
' return a collection of userdetail objects
'-------------------------------------------------------------------------------------
Dim rsUser As Recordset
Dim oDetail As UserDetail
Dim sSQL As String
Dim oCurrentUser As MACROUser
Dim colDetails As Collection
Dim conSecurity As Connection
Dim oLoggedInDetail As UserDetail

    Log "starting GetUserDetails"
    
    If Not ValidateUsername(sUserName) Then
        sMessage = "invalid username"
        Log "invalid username"
        Set GetUsersDetails = Nothing
'EXIT FUNCTION
        Exit Function
    End If
    
    Set oCurrentUser = GetCurrentUser(sSerialisedUser)
    Log "creating security connection"
    Set conSecurity = New Connection
    Call conSecurity.Open(SecurityDatabasePath)
    conSecurity.CursorLocation = adUseClient
    Set colDetails = New Collection
    
    Set oLoggedInDetail = GetOldUserDetails(conSecurity, oCurrentUser.UserName)
    
    sSQL = "SELECT UserName, UserNameFull,Enabled,FailedAttempts,SysAdmin FROM MACROUser"
    
    If Not oLoggedInDetail.SysAdmin Then
        Log "not sys admin"
        sSQL = sSQL & " WHERE SysAdmin = 0"
        If sUserName <> "" Then
            Log "single user: " & sUserName
            sSQL = sSQL & " AND UserName = '" & sUserName & "'"
        End If
    Else
        If sUserName <> "" Then
            Log "single user: " & sUserName
            sSQL = sSQL & " WHERE UserName = '" & sUserName & "'"
        End If
    End If
    
    Set rsUser = New ADODB.Recordset
    rsUser.Open sSQL, conSecurity, adOpenKeyset, , adCmdText
    Log "no of matching users " & rsUser.RecordCount
    Do While Not rsUser.EOF
        'these should never be null
        Set oDetail = New UserDetail
        oDetail.Enabled = (rsUser!Enabled = 1)
        oDetail.FailedAttempts = rsUser!FailedAttempts
        oDetail.UserNameFull = rsUser!UserNameFull
        oDetail.UserName = rsUser!UserName
        oDetail.SysAdmin = (rsUser!SysAdmin = 1)
        colDetails.Add oDetail, oDetail.UserName
        rsUser.MoveNext
    Loop
    rsUser.Close
    Set rsUser = Nothing

    Set GetUsersDetails = colDetails
    
    sMessage = moLog.Messages

End Function


'from frmUserDetails
'-------------------------------------------------------------------------------------
Private Function UpdateUserDetail(oNewDetail As UserDetail, oCurrentUser As MACROUser, _
                            ByRef sFailMessage As String, _
                            conMACRO As Connection, conSecurity As Connection) As Boolean
'-------------------------------------------------------------------------------------
'This updates the user being edited.
'REVISIONS:
'REM 08/12/03 - change password now resets user password to expire if MACRO Setting "expirepassword" = true
'    To tell the user object to expire the password I have passed in "sysadminreset".
'-------------------------------------------------------------------------------------
Dim oSystemMessage As SysMessages
Dim sMessageParameters As String
Dim sHashedPassword As String
Dim sPswdCreateDate As String
Dim sFirstlogin As String
Dim sLastLogin As String
Dim oOldDetail As UserDetail
Dim bPasswordChanged As Boolean
Dim sLog As String
Dim sWhatMessage As String
Dim oLoggedInDetail As UserDetail

    On Error GoTo ErrHandler
    
    Log "Starting UpdateUserDetail"
    
    If Not ValidateUsername(oNewDetail.UserName) Then
        sFailMessage = "invalid username"
        Log "invalid username"
        UpdateUserDetail = False
'EXIT FUNCTION
        Exit Function
    End If
    
    If Not ValidatePassword(oNewDetail.UnEncryptedPassword) Then
        sFailMessage = "invalid new password"
        UpdateUserDetail = False
'EXIT FUNCTION
        Exit Function
    End If
    
    Set oLoggedInDetail = GetOldUserDetails(conSecurity, oCurrentUser.UserName)
    
    'get user details
    Set oOldDetail = GetOldUserDetails(conSecurity, oNewDetail.UserName)
        
    If oOldDetail Is Nothing Then
        'no matching user
        Log "no matching user"
        sFailMessage = "no matching user"
        UpdateUserDetail = False
'EXIT FUNCTION HERE
        Exit Function
    End If
    
    Log "old: " & oOldDetail.ToString
    Log "new: " & oNewDetail.ToString
    
    'check whether there are any changes
    If oNewDetail.UnEncryptedPassword = "" Then
        If oNewDetail.Enabled = oOldDetail.Enabled Then
            If oNewDetail.FailedAttempts = oOldDetail.FailedAttempts Then
                Log "no changes were made"
                sFailMessage = "no changes were made"
                UpdateUserDetail = False
                'EXIT FUNCTION HEReE
                Exit Function
            End If
        End If
    End If
            
    If LCase(oOldDetail.UserName) = "rde" Then
        'exclude rde user
        Log "not allowed to change rde account"
        sFailMessage = "not allowed to change rde account"
        UpdateUserDetail = False
'EXIT FUCNTION HERE
        Exit Function
    End If
    
    
    If Not oLoggedInDetail.SysAdmin And oOldDetail.SysAdmin Then
        'not allowed to change sys admin user if not sys admin
        Log "trying to update SysAdmin user by a new SysAdmin user"
        sFailMessage = "cannot update SysAdmin user when not a SysAdmin user"
        UpdateUserDetail = False
'EXIT FUNCTION HERE
        Exit Function
    End If
    
    If oNewDetail.UnEncryptedPassword <> "" Then
        Log "Password has changed"
        'new password
        'TODO ask Richard about this
        'REM 08/12/03 - Pass in "sysadminreset" to tell user object that this is a sys admin resetting a users password
        sWhatMessage = "sysadminreset"
        
        'if returns false then password change was unsuccessful
        bPasswordChanged = oCurrentUser.ChangeUserPassword(oOldDetail.UserName, oNewDetail.UnEncryptedPassword, sWhatMessage, sHashedPassword, sPswdCreateDate)
        If (GetMACROSetting("expirepassword", "true") = "true") Then
            sFirstlogin = "36000" 'expire new password by setting old date
        Else
            sFirstlogin = LocalNumToStandard(IMedNow)
        End If
        sLastLogin = LocalNumToStandard(IMedNow)
                
        If Not bPasswordChanged Then
            Log "password change failed"
            sFailMessage = sWhatMessage
            'password not changed'Call DialogError(sWhatMessage, gsDIALOG_TITLE)
            Call gLog(oCurrentUser.UserName, gsCHANGE_PSWD, "Change password for user " & oNewDetail.UserName & " failed. " & sWhatMessage, conSecurity)
            'If password fails then all fails
            UpdateUserDetail = False
'EXIT FUNCTION HERE
            Exit Function
        Else
            Log "password change succeeded, unlock account"
            oNewDetail.FailedAttempts = 0
            'when resetting a user password always ensure they are also unlocked
            Call oCurrentUser.gLog(oCurrentUser.UserName, gsCHANGE_PSWD, "The password for user " & oOldDetail.UserName & " was changed using System Management")
            'write New Password to the Message table
                Set oSystemMessage = New SysMessages
                sMessageParameters = oOldDetail.UserName & gsPARAMSEPARATOR & sHashedPassword & gsPARAMSEPARATOR & sLastLogin & gsPARAMSEPARATOR & sFirstlogin & gsPARAMSEPARATOR & sPswdCreateDate
                Call oSystemMessage.AddNewSystemMessage(conMACRO, ExchangeMessageType.PasswordChange, oCurrentUser.UserName, oOldDetail.UserName, "Change Password", sMessageParameters)
                Set oSystemMessage = Nothing
        End If
    End If
    
    If (oOldDetail.Enabled <> oNewDetail.Enabled) Or (oOldDetail.FailedAttempts <> oNewDetail.FailedAttempts) Then
        'enbled status or failed attempts has changed - update MACROUser table
        Log "enabled status or failed attempts has changed - update MACROUser table"
        Call UpdateEnabledFailed(oCurrentUser, conSecurity, oNewDetail)
    End If
    
    If (oNewDetail.FailedAttempts = 0) And (oOldDetail.FailedAttempts <> 0) Then
        'user unlocked - log this
        Log "user unlocked - log this"
        Call gLog(oCurrentUser.UserName, gsUSER_UNLOCKED, "User " & oOldDetail.UserName & " was unlocked", conSecurity)
    End If
    
    'check to see if enabled status has changed if so log the change
    If Not (oOldDetail.Enabled) And oNewDetail.Enabled Then
        'log that user has become enabled
        Log "log that user has become enabled"
        Call gLog(oCurrentUser.UserName, gsUSER_ENABLED, "User " & oOldDetail.UserName & " status changed to enabled", conSecurity)
    End If
    If oOldDetail.Enabled And (Not oNewDetail.Enabled) Then
        'log that user has become disabled
        Log "log that user has become disabled"
        Call gLog(oCurrentUser.UserName, gsUSER_DISABLED, "User " & oOldDetail.UserName & " status changed to disabled", conSecurity)
    End If
    
    'Write edited Users Details to Message table
    Set oSystemMessage = New SysMessages
    With oNewDetail
        Log "Write edited Users Details to Message table"
        sMessageParameters = oOldDetail.UserName & gsPARAMSEPARATOR & .UserNameFull & gsPARAMSEPARATOR & "" & gsPARAMSEPARATOR _
            & -CInt(.Enabled) & gsPARAMSEPARATOR & 0 & gsPARAMSEPARATOR & 0 & gsPARAMSEPARATOR & .FailedAttempts _
            & gsPARAMSEPARATOR & 0 & gsPARAMSEPARATOR & -CInt(.SysAdmin) & gsPARAMSEPARATOR & eUserDetails.udEditUser
    End With
    Call oSystemMessage.AddNewSystemMessage(conMACRO, ExchangeMessageType.User, oCurrentUser.UserName, oOldDetail.UserName, "Edit User Details", sMessageParameters)
    Set oSystemMessage = Nothing
    UpdateUserDetail = True


Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|ChangeUserDetails.UpdateUser"
End Function


'-------------------------------------------------------------------------------------
Private Function GetOldUserDetails(conSecurity As Connection, sUserName As String) As UserDetail
'-------------------------------------------------------------------------------------
'Read single user properties
'-------------------------------------------------------------------------------------
Dim rsUser As Recordset
Dim oDetail As UserDetail
Dim sSQL As String

            sSQL = "SELECT UserName, UserNameFull,Enabled,FailedAttempts,SysAdmin FROM MACROUser" _
                & " WHERE UserName = '" & sUserName & "'"
            Set rsUser = New ADODB.Recordset
            rsUser.Open sSQL, conSecurity, adOpenKeyset, , adCmdText
            If rsUser.RecordCount > 0 Then
                'these should never be null
                Set oDetail = New UserDetail
                oDetail.Enabled = (rsUser!Enabled = 1)
                oDetail.FailedAttempts = rsUser!FailedAttempts
                oDetail.UserNameFull = rsUser!UserNameFull
                oDetail.UserName = rsUser!UserName
                oDetail.SysAdmin = (rsUser!SysAdmin = 1)
            End If
            rsUser.Close
            Set rsUser = Nothing
            Set GetOldUserDetails = oDetail
            

End Function

'-------------------------------------------------------------------------------------
Private Sub UpdateEnabledFailed(oCurrentUser As MACROUser, conSecurity As Connection, oNewDetail As UserDetail)
'-------------------------------------------------------------------------------------
'update date users failed attempts and enabled status
'-------------------------------------------------------------------------------------
Dim sSQL As String

        sSQL = "UPDATE MACROUser SET " _
            & "FailedAttempts=" & oNewDetail.FailedAttempts & ", Enabled=" & -CInt(oNewDetail.Enabled) _
            & " WHERE UserName='" & oNewDetail.UserName & "'"
        conSecurity.Execute sSQL, , adCmdText

End Sub

Private Sub Class_Initialize()
    Set moLog = New APILogger
    moLog.Init
    InitialiseSettingsFile True
    Log "initialising settings file"
End Sub
