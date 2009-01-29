VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LISALogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------
' File: LISALogin.cls
' Copyright: InferMed Ltd 2003 All Rights Reserved
' Author: Nicky Johns/Richard Meinesz, August 2003
' Purpose: Login routine for Interface between MACRO and LISA application
'------------------------------------------------------------------
' REVISIONS
' NCJ 10 Nov 03 - Make sure the "Password expired/about to expire" results are actually returned from LISALogin
'------------------------------------------------------------------

Option Explicit

'--------------------------------------------------------------------------------------------------
Friend Function LISALogin(ByVal sUserName As String, _
                          ByVal sPassword As String, _
                          ByVal sDatabase As String, _
                          ByVal sUserRole As String, _
                          ByRef sMessage As String, _
                          ByRef sUserFullName As Variant, _
                          ByRef sSerialisedUser As String) As LoginResult
'--------------------------------------------------------------------------------------------------
'REM 11/08/03
'Login routine for the LISA application.
'Returns a enum
' NCJ 10 Nov 03 - Code rewritten to make sure the "Password expired/about to expire" results
'               are actually returned
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
    
    Set oUser = New MACROUser
    
    sSecCon = GetSecurityCon
    
    lLoginResult = oUser.Login(sSecCon, sUserName, sPassword, "", "LISA", sMessage)
    
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
            
            'if user does not have LISA UserRole, then fail login
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
        lUserLogin = oUser.Login(sSecCon, sUserName, sPassword, "", "LISA", sMessage, False, sDatabase, sUserRole, True)
    
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
    LISALogin = lLoginResult
    
    Set oUser = Nothing
    
Exit Function
ErrLabel:
    LISALogin = LoginResult.Failed
    sSerialisedUser = ""
    sUserFullName = ""
    sMessage = sMessage & " - Login Error"
End Function

'--------------------------------------------------------------------------------------------------
Private Function GetSecurityCon() As String
'--------------------------------------------------------------------------------------------------
' function returns the connection string for the macro security db specified in the settings file
'--------------------------------------------------------------------------------------------------
    InitialiseSettingsFile True
    GetSecurityCon = GetMACROSetting(MACRO_SETTING_SECPATH, "")
    If GetSecurityCon <> "" Then
        GetSecurityCon = DecryptString(GetSecurityCon)
    End If
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


