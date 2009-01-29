VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APILogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------
' File: APILogin.cls
' Copyright: InferMed Ltd 2004-2008 All Rights Reserved
' Author: Nicky Johns/Richard Meinesz, Feb 2004
' Purpose: Login routines for MACRO Interface (based on original LISA interface)
'------------------------------------------------------------------
' REVISIONS
' NCJ 10 Nov 03 - Make sure the "Password expired/about to expire" results are actually returned from APILogin
' NCJ 2 Feb 04 - This file created from original LISALogin.cls
' NCJ 15 Dec 04 - Removed references to "LISA"
' NCJ 23 Nov 07 - Added Security ConStr to MACROLogin (for ARFG)
' NCJ 26 Feb 08 - Added routines for ARFG Reset Password
' NCJ 3 Mar 08 - Added check for ResetPassword permission
' NCJ 10 Mar 08 - Moved GetSecurityCon and UserExists to modAPIUtils
' NCJ 17 Mar 08 - Removed Security DB parameter from ResetAnotherPassword
'------------------------------------------------------------------

Option Explicit

'--------------------------------------------------------------------------------------------------
Friend Function MACROLogin(ByVal sUserName As String, _
                          ByVal sPassword As String, _
                          ByVal sDatabase As String, _
                          ByVal sUserRole As String, _
                          ByRef sMessage As String, _
                          ByRef sUserFullName As Variant, _
                          ByRef sSerialisedUser As String, _
                          ByRef bGoodSecurity As Boolean, _
                          Optional sSecurityCon As String = "") As LoginResult
'--------------------------------------------------------------------------------------------------
'REM 11/08/03
'Login routine for MACRo API.
'Returns a enum
' NCJ 10 Nov 03 - Code rewritten to make sure the "Password expired/about to expire" results
'               are actually returned
' NCJ 23 Nov 07 - Added bGoodSecurity and optional sSecurityCon argument (Security DB connection string)
'--------------------------------------------------------------------------------------------------
Dim oUser As MACROUser
Dim sSecCon As String
Dim lLoginResult As Long
Dim lUserLogin As Long
Dim bUserRole As Boolean
Dim vUserRole As Variant
Dim sRetMsg As String

    On Error GoTo ErrLabel
    
    ' Initialise return values
    sUserFullName = ""
    sSerialisedUser = ""
    
    ' NCJ 23 Nov 07
    bGoodSecurity = True
    
    If Not ValidateUsername(sUserName) Then
        sMessage = "invalid username"
        MACROLogin = Failed
'EXIT FUNCTION
        Exit Function
    End If
    
    
    If Not ValidatePassword(sPassword) Then
        sMessage = "invalid password"
        MACROLogin = Failed
'EXIT FUNCTION
        Exit Function
    End If
    
    
    If Not ValidateDatabase(sDatabase) Then
        sMessage = "invalid database"
        MACROLogin = Failed
'EXIT FUNCTION
        Exit Function
    End If
    
    
    If Not ValidateRole(sUserRole) Then
        sMessage = "invalid userrole"
        MACROLogin = Failed
'EXIT FUNCTION
        Exit Function
    End If
    
    
    Set oUser = New MACROUser
    
    ' NCJ 23 Nov 07 - Added sSecurityCon
    sSecCon = GetSecurityCon(sSecurityCon)
    ' Does it seem OK?
    bGoodSecurity = (sSecCon <> "")
    If Not bGoodSecurity Then
        sMessage = "invalid security DB"
        MACROLogin = Failed
'EXIT FUNCTION
        Exit Function
    End If

    lLoginResult = oUser.Login(sSecCon, sUserName, sPassword, "", "MACRO API", sMessage)
    
    Select Case lLoginResult
    ' NCJ 10 Nov 03 - If failure or account disabled then don't continue with login
    Case LoginResult.Success, LoginResult.ChangePassword, LoginResult.PasswordExpired
        ' NCJ 10 Nov 03 - Keep hold of password info message
        sRetMsg = sMessage
        ' set up current database so can retrieve the UserRoles
        If oUser.SetCurrentDatabase(sUserName, sDatabase, "", False, True, sMessage) Then
            
            'check if user has been assigned the passed in user role
            bUserRole = False
            For Each vUserRole In oUser.UserRoles
            
                If sUserRole = vUserRole Then
                    bUserRole = True
                End If
            
            Next
            
            'if user does not have correct UserRole, then fail login
            If Not bUserRole Then
                lLoginResult = LoginResult.Failed
                sMessage = "User " & sUserName & " does not have the " & sUserRole & " Role"
            End If
        Else
            ' Can't set up database
            lLoginResult = LoginResult.Failed
        End If
    End Select
    
    ' Only continue for "successful" login
    Select Case lLoginResult
    Case LoginResult.Success, LoginResult.ChangePassword, LoginResult.PasswordExpired
    
        Set oUser = New MACROUser
        
        ' Log in again with DB and user role details
        lUserLogin = oUser.Login(sSecCon, sUserName, sPassword, "", "MACRO API", sMessage, False, sDatabase, sUserRole, True)
    
        Select Case lUserLogin
        Case LoginResult.Success, LoginResult.ChangePassword, LoginResult.PasswordExpired
            sUserFullName = oUser.UserNameFull
            sSerialisedUser = oUser.GetStateHex(False)
            ' NCJ 10 Nov 03 - Return the message from the first login (containing password expiry info)
            If sRetMsg > "" Then
                sMessage = sRetMsg
            End If
        Case Else
            ' We failed for some reason
            sMessage = sMessage & " - Login failed"
            ' This result overrides the earlier password login result
            lLoginResult = lUserLogin
        End Select
    
    End Select
    
    ' NCJ 10 Nov 03 - Return the relevant result
    MACROLogin = lLoginResult
    
    Set oUser = Nothing
    
Exit Function
ErrLabel:
    MACROLogin = LoginResult.Failed
    sSerialisedUser = ""
    sUserFullName = ""
    sMessage = sMessage & " - Login Error - " & Err.Number & ": " & Err.Description
End Function


'--------------------------------------------------------------------------------------------------
Friend Function ChangePasswordRequest(ByRef sSerialisedUser As String, _
                                      ByVal sNewPassword As String, _
                                      ByVal sOldPassword As String, _
                                      ByRef sMessage As String) As Boolean
'--------------------------------------------------------------------------------------------------
'   REM 12/08/2003
'   function changes a user password
'--------------------------------------------------------------------------------------------------
Dim oUser As MACROUser
Dim bResult As Boolean
Dim bOK As Boolean

    Set oUser = New MACROUser
    
    Call oUser.SetStateHex(sSerialisedUser)
    
    Set oUser = ChangePassword(oUser, sOldPassword, sNewPassword, True, bResult, sMessage)
    If bResult Then
        bOK = True
        sSerialisedUser = oUser.GetStateHex(False)
    Else
        bOK = False
    End If
    Set oUser = Nothing

    ChangePasswordRequest = bOK
    
End Function

'--------------------------------------------------------------------------------------------------
Private Function ChangePassword(ByRef oUser As MACROUser, _
                                ByVal sOldPassword As String, _
                                ByVal sNewPassword As String, _
                                ByVal bCheckPassword As Boolean, _
                                ByRef bResult As Boolean, _
                                ByRef sMessage As String) As MACROUser
'--------------------------------------------------------------------------------------------------
'   REM 12/08/2003
'   function changes a user password
'--------------------------------------------------------------------------------------------------
Dim eLogin As LoginResult
    
    'if requested, check old password
    If bCheckPassword Then
        eLogin = oUser.Login(GetSecurityCon(), oUser.UserName, sOldPassword, "", "MACRO Web Data Entry", "", True)
    Else
        eLogin = LoginResult.Success
    End If
    
    If (eLogin <> LoginResult.Failed) Then
        'attempt update password
        bResult = oUser.ChangeUserPassword(oUser.UserName, sNewPassword, sMessage, "", "")

    Else
        bResult = False
    End If
    
    Set ChangePassword = oUser
    
End Function

'----------------------------------------------------------------------------------------'
Public Function ResetAnotherPassword(ByRef sSerialisedUser As String, _
                                    ByVal sUserName As String, _
                                    ByVal sNewPwd As String, _
                                    ByRef sMessage As String) As Long
'----------------------------------------------------------------------------------------'
' NCJ 25 Feb 08 - Reset another user's password
' sSecurity can be "" to use default Security DB
' Returns 0 if success
' Returns 1 if user name exists but account is disabled
' Returns 2 if user name does not exist
' Returns 3 if invalid password
' Returns 4 if no ResetPassword permission
'----------------------------------------------------------------------------------------'
Dim oCurrentUser As MACROUser
Dim bPwdOk As Boolean
Dim sHash As String
Dim sDate As String
Dim sSec As String

    On Error GoTo ErrLabel

    sMessage = ""
    
    ' DB connection is valid
    Set oCurrentUser = New MACROUser
    Call oCurrentUser.SetStateHex(sSerialisedUser)
    ' Get the Security DB
    sSec = DecryptString(oCurrentUser.SecurityDBConEncrypted)
    ' Check the logged in user has permission to reset password
    If Not oCurrentUser.CheckPermission(gsFnResetPassword) Then
        ' User does not have permission to Reset Password
        ResetAnotherPassword = 4
        sMessage = "Current user does not have permission to reset passwords"
    Else
        ' Check the target user exists and is enabled
        Select Case UserExists(sSec, sUserName)
        Case 0
            ' User name OK - go ahead and reset password
            ' Signify a Sys Admin Reset (to force a change on next login)
            sMessage = "sysadminreset"
            bPwdOk = oCurrentUser.ChangeUserPassword(sUserName, sNewPwd, sMessage, sHash, sDate)
            If bPwdOk Then
                ' Result = Success
                ResetAnotherPassword = 0
                sMessage = ""
            Else
                ' Result = Invalid password (sMessage already set up)
                ResetAnotherPassword = 3
            End If
        Case 1
            ResetAnotherPassword = 1
            sMessage = "User account is disabled"
        Case 2
            ResetAnotherPassword = 2
            sMessage = "User name not found"
        Case Else
            ' Error
            ResetAnotherPassword = -1
        End Select
    End If
    Set oCurrentUser = Nothing
    
Exit Function
ErrLabel:
    ResetAnotherPassword = -1
End Function
