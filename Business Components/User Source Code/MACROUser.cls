VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MACROUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       MACROUser.cls
'   Author:     Richard Meinesz, September 2002
'   Purpose:
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
'   Revisions:
'TA 7/11/2002: Made class serialisable. NB - only certain properties are serialised - used for the web
'               basically permissions ,permitted users and sites - and info for data lists
'TA 11/02/2002: security connection string added to serialisation
' NCJ 4 Mar 03 - Added access to Batch Validation module
' ic 23/06/2003 added bLogSuccess param, allows www to only log when login is complete: Login()
' MLM 03/07/03: Pass connection string to LoadUserRoleFunctions
' NCJ 27 Nov 03 - Ensure we close everything down in the Terminate event
' REM 17/12/03 - Added check for AllStudies/AllSites permission in LoadStudySitePermissions routine
'               so in Sites and Studies objects only do IN statment if not AllStudies/AllSites
' REM 05/02/04 - In Login routine added message to log saying user does not have permission to enter module (only for silent login)
' DPH 24/03/2004 - Clean up collections in class_terminate, tidy other objects
' NCJ 12 March 2008 - Added SecurityDBConEncrypted property to support the API
'----------------------------------------------------------------------------------------'

Option Explicit

Public Enum LoginResult
    Success = 0
    AccountDisabled = 1
    Failed = 2
    ChangePassword = 3
    PasswordExpired = 4
End Enum

Private moPassword As Password
Private moDatabase As Database
Private moPasswordPolicy As PasswordPolicy
Private mcolDatabases As Collection

Private moPermissions As Permissions

Private msUserName As String
Private msUserNameFull As String

Private msDatabaseCode As String
Private moStudies As Studies
Private moSites As Sites
Private mvStudySite As Variant
Private msSecurityDBSubVersion As String
Private msSecurityDBVersion As String

'store the object user to get lists
Private moDataLists As UserDataLists

'security connection
Private mconSecurity As Connection
'store the security connection
Private msSecurityCon As String

Private mconMACRO As Connection
Private msCurrentDBConnection As String


Private msAppTitle As String

Private mbLoggedIn As Boolean

'serialised version of study site combos
Private msStudySite As String

Public UserSettings As UserSettings

Private msDBSiteCode As String

'----------------------------------------------------------------------------------------'
Public Property Get SecurityDatabaseType() As eMACRODatabaseType
'----------------------------------------------------------------------------------------'
    
    Select Case Connection_Property(CONNECTION_PROVIDER, msSecurityCon)
    
    Case CONNECTION_MSDAORA, CONNECTION_ORAOLEDB_ORACLE
        SecurityDatabaseType = mdtOracle80
    Case CONNECTION_SQLOLEDB
        SecurityDatabaseType = mdtSQLServer
    End Select
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get DBSiteCode()
'----------------------------------------------------------------------------------------'
'this will only work once you are logged in
'----------------------------------------------------------------------------------------'
    DBSiteCode = msDBSiteCode
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get DBIsServer() As Boolean
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    DBIsServer = (msDBSiteCode = "")
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get CurrentDBConString() As String
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    CurrentDBConString = msCurrentDBConnection

End Property

'----------------------------------------------------------------------------------------'
Public Property Get Studies() As Studies
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Set Studies = moStudies
    
End Property

'----------------------------------------------------------------------------------------'
Friend Property Get SecurityConnectionString() As String
'----------------------------------------------------------------------------------------'

    SecurityConnectionString = msSecurityCon
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get SecurityDBConEncrypted() As String
'----------------------------------------------------------------------------------------'
' NCJ 12 March 08 - Return Security DB connection string (encrypted)
' Needed by the API
'----------------------------------------------------------------------------------------'

    SecurityDBConEncrypted = EncryptString(msSecurityCon)

End Property

'----------------------------------------------------------------------------------------'
Public Property Get UserName() As String
'----------------------------------------------------------------------------------------'

    UserName = msUserName
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get UserNameFull() As String
'----------------------------------------------------------------------------------------'

    UserNameFull = msUserNameFull
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get DatabaseCode() As String
'----------------------------------------------------------------------------------------'

    DatabaseCode = msDatabaseCode

End Property

'----------------------------------------------------------------------------------------'
Public Property Get Database() As Database
'----------------------------------------------------------------------------------------'

    Set Database = moDatabase

End Property

'----------------------------------------------------------------------------------------'
Public Property Get PasswordPolicy() As PasswordPolicy
'----------------------------------------------------------------------------------------'

    Set PasswordPolicy = moPasswordPolicy

End Property

'----------------------------------------------------------------------------------------'
Public Property Get SysAdmin() As Boolean
'----------------------------------------------------------------------------------------'
    
    SysAdmin = (moPassword.SysAdmin = 1)
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get SecurityDBVersion() As String
'----------------------------------------------------------------------------------------'

    SecurityDBVersion = msSecurityDBVersion

End Property

'----------------------------------------------------------------------------------------'
Public Property Get SecurityDBSubVersion() As String
'----------------------------------------------------------------------------------------'

    SecurityDBSubVersion = msSecurityDBSubVersion

End Property

'----------------------------------------------------------------------------------------'
Public Function SetCurrentDatabase(sUserName As String, ByVal sDatabase As String, sDefaultHTMLLocation As String, _
                                   bSetConnection As Boolean, bLoadUserRoles As Boolean, _
                                   ByRef sMessage As String) As Boolean
'----------------------------------------------------------------------------------------'
'REM 19/09/02
'Loads all the selected database's properties and creats the connection
'----------------------------------------------------------------------------------------'
Dim sConnection As String
Dim bConnection As Boolean
Dim sDbType As String
Dim bLoadDb As Boolean

    On Error GoTo ErrLabel

    'load all the properties for the selected database
    bLoadDb = moDatabase.Load(mconSecurity, sUserName, sDatabase, sDefaultHTMLLocation, bLoadUserRoles, sMessage)
    
    Set UserSettings = New UserSettings
    
    If bLoadDb Then
        'set user database selection
        msDatabaseCode = sDatabase
        
        bConnection = True
         'get the connection string
        sConnection = moDatabase.ConnectionString
        If bSetConnection Then

            'craete the connection
            bConnection = CreateDBConnection(sConnection, sMessage)
            
        End If
        SetCurrentDatabase = bConnection
    Else
        SetCurrentDatabase = False
    End If
    
    msCurrentDBConnection = sConnection

    Call UserSettings.Load(sConnection, msUserName)
        
Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "UserLogin.SetCurrentDatabase"
End Function

'----------------------------------------------------------------------------------------'
Private Function CreateDBConnection(sConnection As String, ByRef sMessage As String) As Boolean
'----------------------------------------------------------------------------------------'
'REM 30/10/02
'Create the MACRO database connection
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsSiteCode As Recordset

    On Error GoTo ErrLabel

    'create the connection
    Set mconMACRO = New ADODB.Connection
    mconMACRO.Open sConnection
    mconMACRO.CursorLocation = adUseClient
    
    
    sSQL = "SELECT SettingValue FROM MACRODBSetting" _
        & " WHERE SettingSection = 'datatransfer'" _
        & " AND SettingKey = 'dbsitename'"
    Set rsSiteCode = New ADODB.Recordset
    rsSiteCode.Open sSQL, mconMACRO, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsSiteCode.RecordCount > 0 Then
        msDBSiteCode = rsSiteCode!SettingValue
    Else
        msDBSiteCode = ""
    End If
    
    rsSiteCode.Close
    Set rsSiteCode = Nothing
    
    CreateDBConnection = True
        
Exit Function
ErrLabel:
    sMessage = Err.Description
    CreateDBConnection = False
End Function

'----------------------------------------------------------------------------------------'
Public Property Get UserDatabases() As Collection
'----------------------------------------------------------------------------------------'
'REM 19/09/02
'Returns a collection of user database codes
'----------------------------------------------------------------------------------------'

    Set UserDatabases = mcolDatabases

End Property

'----------------------------------------------------------------------------------------'
Public Property Get UserRoles() As Collection
'----------------------------------------------------------------------------------------'
'REM 20/09/02
'Returns a collection of user role codes for a selected database
'----------------------------------------------------------------------------------------'

    Set UserRoles = moDatabase.UserRoles

End Property

'----------------------------------------------------------------------------------------'
Public Sub SetUserRole(ByVal sRoleCode As String)
'----------------------------------------------------------------------------------------'
'REM 20/09/02
'Loads all the permissions (function codes) for a specific role code and study site permissions
'----------------------------------------------------------------------------------------'

    'MLM 03/07/03:
    Call moPermissions.LoadUserRoleFunctions(mconSecurity, sRoleCode, msSecurityCon)
    'Loads all the study/site combinations that the user has under the given rolecode
    Call LoadStudySitePermissions(msUserName, sRoleCode, False)
    
End Sub

'----------------------------------------------------------------------------------------'
Public Property Get UserRole() As String
'----------------------------------------------------------------------------------------'

    UserRole = moPermissions.UserRole

End Property

'----------------------------------------------------------------------------------------'
Public Function CheckPermission(sPermission As String) As Boolean
'----------------------------------------------------------------------------------------'
'REM 23/09/02
'Checks if a user has specific permission
'----------------------------------------------------------------------------------------'
    
    CheckPermission = moPermissions.CheckPermission(sPermission)

End Function

'----------------------------------------------------------------------------------------'
Public Function SilentLogin(sSecurityCon As String, sDatabase As String, sDefaultHTMLLocation As String, _
                            ByRef sMessage As String) As Boolean
'----------------------------------------------------------------------------------------'
'REM 10/10/02
'Silent login that requires no username or password
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrLabel

    'create and open a new security connection
    Set mconSecurity = New Connection
    mconSecurity.Open sSecurityCon
    mconSecurity.CursorLocation = adUseClient
    
    Set moDatabase = New Database
    
    SilentLogin = SetCurrentDatabase("", sDatabase, sDefaultHTMLLocation, True, False, sMessage)
    
    msUserName = "AutoImport"
    
Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "UserLogin.SilentLogin"
End Function


'----------------------------------------------------------------------------------------'
Public Function LoginAD(sSecurityCon As String, ByRef sUserName As String, sPassword As String, _
                      sDefaultHTMLLocation As String, sAppTitle As String, _
                      ByRef sMessage As String, Optional bCheckDetailsOnly As Boolean = False, _
                      Optional sDatabase As String = "", Optional sRoleCode As String = "", _
                      Optional bLogSuccess As Boolean = True) As LoginResult
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
Dim oDatabases As Databases
Dim sSQL As String
Dim rsVersion As ADODB.Recordset
Dim oADLogin As IMEDActiveDirectory.ADLogin


    On Error GoTo ErrLabel
    
    'create and open a new security connection
    Set mconSecurity = New Connection

    'store security connection string
    msSecurityCon = sSecurityCon

    mconSecurity.Open sSecurityCon
    mconSecurity.CursorLocation = adUseClient
    
    
    Set oADLogin = New IMEDActiveDirectory.ADLogin
    Call oADLogin.Init(sSecurityCon)
    
    'perform AD login without application check
'    If (oADLogin.Login("", sMessage)) Then
    'perform AD login with application check
    If (oADLogin.Login("MAC", sMessage)) Then
        LoginAD = LoginResult.Success

        'get the security database version and sub version
        sSQL = "SELECT MACROVersion, BuildSubVersion FROM SecurityControl"
        Set rsVersion = New ADODB.Recordset
        rsVersion.Open sSQL, mconSecurity, adOpenKeyset, adLockReadOnly, adCmdText
    
        'set the version property
        msSecurityDBVersion = rsVersion![MACROVersion]
        'set the Subversion property
        msSecurityDBSubVersion = rsVersion![BuildSubVersion]
    
        ' close recordset & tidy
        rsVersion.Close
        Set rsVersion = Nothing
    
        'Set modular level password object
        Set moPassword = New Password
    
        sMessage = ""
    
        msAppTitle = sAppTitle


        'initialise the password object
        Call moPassword.UserExists(mconSecurity, msSecurityCon, sUserName, sMessage)

        'set msUserName to sUsername here as passwordexists could correct case
        msUserName = sUserName
        msUserNameFull = moPassword.UserNameFull


        LoginAD = LoginResult.Success
        If bLogSuccess Then
            Call gLog(sUserName, "LoginAD", "Login success: " & msAppTitle)
        End If


        Set moPermissions = New Permissions
        Set moDatabase = New Database

        'setup and load password policy
        Set moPasswordPolicy = New PasswordPolicy
        Call moPasswordPolicy.Load(mconSecurity)

        'silent login, if Database and role code are supplied then set-up database and role
        If (sDatabase <> "") And (sRoleCode <> "") Then
            'loads all the properties for the database
            Call SetCurrentDatabase(sUserName, sDatabase, sDefaultHTMLLocation, True, True, sMessage)
            'Loads all the permissions for the rolecode
            Call SetUserRole(sRoleCode)

        Else 'normal login and user will select database and rolecode
            Set oDatabases = New Databases

            'return all the user databases and return the code for the first one
            Set mcolDatabases = oDatabases.UserDatabases(mconSecurity, sUserName)
        End If
        
        ' tidy up
        Set oDatabases = Nothing
        
    Else
        LoginAD = LoginResult.Failed
        Call gLog(sUserName, "LoginAD", "Active Directory Login Failed: " & sMessage)
    End If
    

    mbLoggedIn = (LoginAD = Success)
    
Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "UserLogin.LoginAD"

End Function

'----------------------------------------------------------------------------------------'
Public Function Login(sSecurityCon As String, ByRef sUserName As String, sPassword As String, _
                      sDefaultHTMLLocation As String, sAppTitle As String, _
                      ByRef sMessage As String, Optional bCheckDetailsOnly As Boolean = False, _
                      Optional sDatabase As String = "", Optional sRoleCode As String = "", _
                      Optional bLogSuccess As Boolean = True) As LoginResult
'----------------------------------------------------------------------------------------'
'REM 17/09/02
'Checks users exists, if password is correct and if password is still valid
'Return value indicates if login was successful, failed, password expired or if password needs changing soon
'TA 11/11/2002: If blank password is passed in then username is assumed to be username and password
'ic 23/06/2003 added bLogSuccess param, allows www to only log when login is complete
'REM 05/02/04 - Added message to log saying user does not have permission to enter module
'DPH 23/03/2004 - close recordsets / tidy objects
'----------------------------------------------------------------------------------------'
Dim lValidDaysLeft As Long
Dim oDatabases As Databases
'Dim oPermissions As Permissions
Dim sDatabaseCode As String
Dim sPasswordAndDate As String
Dim sHashedPasswordAndDate As String
Dim sSQL As String
Dim rsVersion As ADODB.Recordset
Dim bSerialisedInput As Boolean
Dim sTemp As String
Dim sMsg As String

    If (sPassword = "") Then
        bSerialisedInput = True
        sTemp = DecryptString(sUserName)
        sUserName = Split(sTemp, ",")(0)
        sPassword = Split(sTemp, ",")(1)
    End If

    On Error GoTo ErrLabel

    'create and open a new security connection
    Set mconSecurity = New Connection

    'store security connection string
    msSecurityCon = sSecurityCon

    mconSecurity.Open sSecurityCon
    mconSecurity.CursorLocation = adUseClient

    'get the security database version and sub version
    sSQL = "SELECT MACROVersion, BuildSubVersion FROM SecurityControl"
    Set rsVersion = New ADODB.Recordset
    rsVersion.Open sSQL, mconSecurity, adOpenKeyset, adLockReadOnly, adCmdText

    'set the version property
    msSecurityDBVersion = rsVersion![MACROVersion]
    'set the Subversion property
    msSecurityDBSubVersion = rsVersion![BuildSubVersion]

    ' close recordset & tidy
    rsVersion.Close
    Set rsVersion = Nothing

    'Set modular level password object
    Set moPassword = New Password

    sMessage = ""

    msAppTitle = sAppTitle

    'check if UserName exists, this initialises the password object
    If moPassword.UserExists(mconSecurity, msSecurityCon, sUserName, sMessage) Then

        'set msUserName to sUsername here as passwordexists could correct case
        msUserName = sUserName

        msUserNameFull = moPassword.UserNameFull

        If bSerialisedInput Then
                sHashedPasswordAndDate = sPassword
        Else
            'add password create date to the password
             sPasswordAndDate = sPassword & LocalNumToStandard(CStr(moPassword.PasswordCreated))
             'hash the password and create date string
             sHashedPasswordAndDate = HashHexEncodeString(sPasswordAndDate)

        End If

        'check password is correct and account is not locked out or disabled
        Select Case moPassword.PasswordCheck(mconSecurity, sHashedPasswordAndDate, sMessage)

        Case ePasswordCheck.pcPass
            If moPassword.PasswordExpiryCheck(sUserName, lValidDaysLeft) Then
                Select Case lValidDaysLeft
                    Case 2 To 7
                        sMessage = "You have " & lValidDaysLeft & " days before your password expires." _
                                & vbCr & "Do you wish to change it now?"
                        Login = LoginResult.ChangePassword

                    Case 1
                        sMessage = "Your password will expire after today - do you wish to change it now?"
                        Login = LoginResult.ChangePassword

                    Case Is < 1
                        ' Password has expired - they must change it now
                        sMessage = "Your password has expired, and it must be changed now."
                        Login = LoginResult.PasswordExpired

                    Case Else
                        Login = LoginResult.Success
                        'ic 23/06/2003 optional log condition
                        If bLogSuccess Then
                            Call gLog(sUserName, "Login", "Login success: " & msAppTitle)
                        End If
                End Select
            Else 'didn't require password expiry check, so login success
                Login = LoginResult.Success
                'ic 23/06/2003 optional log condition
                If bLogSuccess Then
                    Call gLog(sUserName, "Login", "Login success: " & msAppTitle)
                End If
            End If

        Case ePasswordCheck.pcFailed

            Login = LoginResult.Failed
            Call gLog(sUserName, "Login", "Login Failed, User Password incorrect: " & msAppTitle)

        Case ePasswordCheck.pcAccountDisabled

            Login = LoginResult.AccountDisabled
            Call gLog(sUserName, "Login", "Login Failed, Account disabled: " & msAppTitle)

        Case ePasswordCheck.pcAccountLockout

            Login = LoginResult.AccountDisabled
            Call gLog(sUserName, "Login", "Login Failed, Account locked out: " & msAppTitle)

        End Select

    Else
        Login = LoginResult.Failed
        Call gLog(sUserName, "Login", "Login Failed, User Name does not exist: " & msAppTitle)
    End If

    ' this may be a timeout login so do not do all initialization again if it is
    If bCheckDetailsOnly = False Then
        'if login doesn't fail then initalise the user databases object and user permission object
        If (Login = LoginResult.Success) Or (Login = LoginResult.PasswordExpired) Or (Login = LoginResult.ChangePassword) Then

            Set moPermissions = New Permissions
            Set moDatabase = New Database

            'setup and load password policy
            Set moPasswordPolicy = New PasswordPolicy
            Call moPasswordPolicy.Load(mconSecurity)

            'silent login, if Database and role code are supplied then set-up database and role
            If (sDatabase <> "") And (sRoleCode <> "") Then
                'loads all the properties for the database
                Call SetCurrentDatabase(sUserName, sDatabase, sDefaultHTMLLocation, True, True, sMessage)
                'Loads all the permissions for the rolecode
                Call SetUserRole(sRoleCode)

                'check user has permission to enter module
                If UserHasAccessToModule(sAppTitle, sMessage) Then
                    Login = LoginResult.Success
                Else
                    Login = LoginResult.Failed
                    'REM 05/02/04 - Added message to log saying user does not have permission to enter module
                    Call gLog(sUserName, "Login", "User did not have permission to access " & msAppTitle)
                End If

            Else 'normal login and user will select database and rolecode
                Set oDatabases = New Databases

                'return all the user databases and return the code for the first one
                Set mcolDatabases = oDatabases.UserDatabases(mconSecurity, sUserName)
                'set up the user roles for a selected database (uses the first database in the collection above, which is returned ByVal)
                'Call SetCurrentDatabase(sUserName, sDatabaseCode, sDefaultHTMLLocation, False, True, sMsg)
'                If sMsg <> "" Then
'                    sMessage = sMsg
'                End If
            End If
        End If
    End If

    ' tidy up
    Set oDatabases = Nothing

    mbLoggedIn = (Login = Success)
    
Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "UserLogin.Login"
End Function

'----------------------------------------------------------------------------------------'
Public Function AllUserDatabases(sSecurityCon As String, ByVal sUserName As String, _
                                sDefaultHTMLLocation As String) As Collection
'----------------------------------------------------------------------------------------'
'REM 09/10/02
'Returns a collection of database objects
'----------------------------------------------------------------------------------------'
Dim oDatabases As Databases
Dim oDatabase As Database
Dim colUserDatabases As Collection
Dim colDatabases As Collection
Dim conSecurity As Connection
Dim sDatabase As String
Dim i As Integer
Dim sMessage As String
Dim bLoad As Boolean

    Set oDatabases = New Databases
    Set colDatabases = New Collection
    Set colUserDatabases = New Collection
    
    'create and open a new security connection
    Set conSecurity = New Connection
    conSecurity.Open sSecurityCon
    
    'get collection of user databases
    Set colDatabases = oDatabases.UserDatabases(mconSecurity, sUserName)
    
    'loop through collection
    For i = 1 To colDatabases.Count
    
        Set oDatabase = New Database
        
        sDatabase = colDatabases(i)
        
        'load database object
        bLoad = oDatabase.Load(conSecurity, sUserName, sDatabase, sDefaultHTMLLocation, True, sMessage)
        
        'add database object to collection
        colUserDatabases.Add oDatabase
        
    Next
    
    Set AllUserDatabases = colUserDatabases
    
    
End Function

'----------------------------------------------------------------------------------------'
Public Function LogOff()
'----------------------------------------------------------------------------------------'

    Call gLog(msUserName, "LogOff", "User logged off: " & msAppTitle)

    ADO_CloseConnection mconMACRO
    ADO_CloseConnection mconSecurity

    
    mbLoggedIn = False
    
End Function

'----------------------------------------------------------------------------------------'
Public Sub SetToNotLoggedIn()
'----------------------------------------------------------------------------------------'
' this sets trhe user to logged out without destroying all the internal state
'----------------------------------------------------------------------------------------'

    mbLoggedIn = False
    
End Sub

'----------------------------------------------------------------------------------------'
Public Property Get LoggedIn() As Boolean
'----------------------------------------------------------------------------------------'
' this sets trhe user to logged out without destroying all the internal state
'----------------------------------------------------------------------------------------'

    LoggedIn = mbLoggedIn
    
End Property

'----------------------------------------------------------------------------------------'
Public Function CreateNewUser(sNewUserName As String, sNewUserNameFull As String, sPassword As String, _
                              bEnabled As Boolean, bSysAdmin As Boolean, ByRef sEncryptedPasswordAndDate As String, _
                              ByRef sPasswordCreated As String, ByRef sMessage As String, Optional bActiveDirectory As Boolean) As Boolean
'----------------------------------------------------------------------------------------'
'REM 24/10/02
'Used to create a new user
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim sFirstLogin As String
Dim sLastlogin As String
Dim nFailedAttempts As Integer
Dim sDefaultUserRole As String
Dim bSucceed As Boolean
Dim nEnabled As Integer
Dim nSysAdmin As Integer
Dim sPasswordAndCreateDate As String
Dim nActiveDirectory As Integer
    
    On Error GoTo ErrLabel
    
    'ic 20/12/2005 move this to the argument list when you dont mind breaking compatibility
    If IsMissing(bActiveDirectory) Then bActiveDirectory = False
    
    'Initialise the Settings File
    InitialiseSettingsFile True
    
    If (GetMACROSetting("expirepassword", "true") = "true") Then
        sFirstLogin = "36000" ' old date to force the password to expire
    Else
        sFirstLogin = SQLStandardNow
    End If
    
    sLastlogin = SQLStandardNow
    sPasswordCreated = SQLStandardNow
    nFailedAttempts = 0
    sDefaultUserRole = " "
    nEnabled = -CInt(bEnabled)
    nSysAdmin = -CInt(bSysAdmin)
    nActiveDirectory = -CInt(bActiveDirectory)
    
    'call changepassowrd with the optional last parameter set to false so it just checks the password against the policy
    bSucceed = moPassword.ChangePassword(mconSecurity, sNewUserName, sPassword, sMessage, , , False)
    
    'If the checks pass then create the new user in the secuirty database
    If bSucceed Then
        'encrypt the password with the create date
        sPasswordAndCreateDate = sPassword & sPasswordCreated
        sEncryptedPasswordAndDate = HashHexEncodeString(sPasswordAndCreateDate)
        'do insert
        sSQL = "INSERT INTO MacroUser " _
            & "(UserName,UserNameFull,UserPassword,Enabled,Lastlogin,FirstLogin,DefaultUserRoleCode, " _
            & " FailedAttempts,PasswordCreated,SysAdmin"
        If (bActiveDirectory) Then
            sSQL = sSQL & ",Authentication"
        End If
        sSQL = sSQL & ")" _
            & " VALUES ('" & sNewUserName & "','" & ReplaceQuotes(sNewUserNameFull) & "','" & sEncryptedPasswordAndDate & "'," & nEnabled & "," & sLastlogin & "," _
            & sFirstLogin & ",'" & sDefaultUserRole & "'," & nFailedAttempts & "," & sPasswordCreated & "," & nSysAdmin
        If (bActiveDirectory) Then
            sSQL = sSQL & "," & nActiveDirectory
        End If
        sSQL = sSQL & ")"
        mconSecurity.Execute sSQL, , adCmdText
        
    End If
    
    'return
    CreateNewUser = bSucceed

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "UserLogin.CreateNewUser"
End Function

'----------------------------------------------------------------------------------------'
Public Function ChangeUserPassword(ByVal sUserName As String, ByVal sNewPassword As String, _
                                   ByRef sMessage As String, ByRef sHashedPassword As String, _
                                   ByRef sCreateDate As String) As Boolean
'----------------------------------------------------------------------------------------'
'REM 18/09/02
'Changes user password
'----------------------------------------------------------------------------------------'
Dim bPasswordChanged As Boolean

    On Error GoTo ErrLabel

    If mconSecurity Is Nothing Then
        'we have no connecition
        If msSecurityCon <> "" Then
            'we have security con string - open it
            Set mconSecurity = New Connection
            mconSecurity.Open msSecurityCon
        End If
    End If
    
    bPasswordChanged = moPassword.ChangePassword(mconSecurity, sUserName, sNewPassword, sMessage, sHashedPassword, sCreateDate, True)
    
'TA 13/03/2003: this is already changed to the encrypted version within moPassword.ChangePassword
'    'if passowrd change is successful then change password in the password class
'    If bPasswordChanged Then
'        moPassword.Password = sNewPassword
'    End If

    ChangeUserPassword = bPasswordChanged

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "UserLogin.ChangeUserPassword"
End Function

'------------------------------------------------------------------------------------'
Public Function UserHasAccessToModule(sApplicationTitle As String, ByRef sMessage As String) As Boolean
'------------------------------------------------------------------------------------'
' Returns TRUE if user's role allows access to this application
' Assume current user's function collection already set up
'REM 13/01/03 - Pass in ApplicationTitle parameter as was getting incorrect App.Title
'REM 14/02/03 - Changed to use new Application constants
' NCJ 4 Mar 03 - Added Batch Validation
'------------------------------------------------------------------------------------'
Dim sAppTitle As String

    On Error GoTo ErrLabel
    
    ' This gets filled with app title if they don't have access
    sAppTitle = ""
    
    Select Case sApplicationTitle
    
    Case APPTITLE_SM
        If CheckPermission(gsFnSystemManagement) = False Then
            sAppTitle = APPTITLE_SM
        End If
    'REM 14/02/03 - MACRO Exchange no longer exists
'    Case "MACRO_EX"
'        If CheckPermission(gsFnExchange) = False Then
'            sAppTitle = "MACRO Exchange"
'        End If
        
    Case APPTITLE_DR
            If CheckPermission(gsFnDataReview) = False Then
                sAppTitle = APPTITLE_DR
            End If
            
    Case APPTITLE_DE
            If CheckPermission(gsFnDataEntry) = False Then
                sAppTitle = APPTITLE_DE
            End If
        
    Case APPTITLE_LM
            If CheckPermission(gsFnLibraryManagement) = False Then
                sAppTitle = APPTITLE_LM
            End If
            
    Case APPTITLE_SD
            If CheckPermission(gsFnStudyDefinition) = False Then
                sAppTitle = APPTITLE_SD
            End If

    Case APPTITLE_DV
        If Not CheckPermission(gsFnCreateDataViews) Then
            sAppTitle = APPTITLE_DV
        End If

    Case APPTITLE_QM
        If CheckPermission(gsFnQueryModule) = False Then
            sAppTitle = APPTITLE_QM
        End If
        
    Case APPTITLE_BD
        If CheckPermission(gsFnBatchDataEntry) = False Then
            sAppTitle = APPTITLE_BD
        End If

    ' NCJ 4 Mar 03
    Case APPTITLE_BV
        If CheckPermission(gsFnBatchValidation) = False Then
            sAppTitle = APPTITLE_BV
        End If
    End Select

    If sAppTitle > "" Then  ' They don't have access
        sMessage = "You do not have permission to access " & sAppTitle
        UserHasAccessToModule = False
    Else
        sMessage = ""
        UserHasAccessToModule = True
        ' Let frmMenu enable its menu options etc.
        ' Need definition of CheckUserRights in each frmMenu
'***UNCOMMENT***        'Call frmMenu.CheckUserRights
    
    End If
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "UserLogin.UserHasAccessToModule"

End Function

'----------------------------------------------------------------------------------------'
Public Sub ReloadStudySitePermissions()
'----------------------------------------------------------------------------------------'
'REM 21/05/03
'Routine to reload study-site permissions when ever they are changed in system management
'----------------------------------------------------------------------------------------'
    
    'MLM 03/07/03:
    Call moPermissions.LoadUserRoleFunctions(mconSecurity, moPermissions.UserRole, msSecurityCon)
    Call LoadStudySitePermissions(msUserName, moPermissions.UserRole, False)

End Sub

'----------------------------------------------------------------------------------------'
Private Function LoadStudySitePermissions(ByVal sUserName As String, ByVal sRoleCode As String, _
                                            bNoDBAccess As Boolean) As Boolean
'----------------------------------------------------------------------------------------'
'REM 26/09/02
'Load all the study site combonations for a user and role in the current database
'TA 7/11/2002: If bNoDBAccess then the datbase is not accessed becasue mvStudySite, mostudies and moSites already exist
'TA 06/01/2003: Added UserDataLists to SerialiizedUser
'REM 17/12/03 - Added routine to check if user has AllStudies/AllSites permission, so can pass this into Site and Study load
'----------------------------------------------------------------------------------------'
Dim rsStudySite As ADODB.Recordset
'Dim vStudySite As Variant
Dim i As Integer
Dim sSQL As String
Dim sSites As String
Dim sStudies As String
Dim bAllStudiesAllSites As Boolean

    On Error GoTo ErrLabel

    If Not bNoDBAccess Then
        'we need to load from database
        'Select all the specific Study-Site combinations for a user for a specific role
        sSQL = "SELECT Clinicaltrial.ClinicalTrialId, TrialSite.TrialSite, ClinicalTrial.ClinicalTrialName" _
            & " FROM UserRole, TrialSite, ClinicalTrial" _
            & " WHERE (TrialSite.TrialSite = UserRole.SiteCode OR UserRole.SiteCode = 'AllSites')" _
            & " AND (ClinicalTrial.ClinicalTrialName = UserRole.StudyCode OR UserRole.StudyCode = 'AllStudies')" _
            & " AND ClinicalTrial.ClinicalTrialId = TrialSite.ClinicalTrialId" _
            & " AND UserRole.UserName = '" & sUserName & "'" _
            & " AND UserRole.RoleCode = '" & sRoleCode & "'" _
            & " ORDER BY ClinicalTrial.ClinicalTrialName,TrialSite.TrialSite"
    
        Set rsStudySite = New ADODB.Recordset
        rsStudySite.Open sSQL, mconMACRO, adOpenKeyset, adLockPessimistic, adCmdText
        msStudySite = ADO_SerialiseRecordset(rsStudySite)
    Else
       Set rsStudySite = ADO_DeSerialiseRecordset(msStudySite)
    End If
    
    If rsStudySite.RecordCount > 0 Then
        'convert recordset into an array
        mvStudySite = rsStudySite.GetRows
   
        rsStudySite.Close
        Set rsStudySite = Nothing
        'read all studies into a comma delimited string
        sStudies = mvStudySite(0, 0)
        For i = 1 To UBound(mvStudySite, 2)
            sStudies = sStudies & "," & mvStudySite(0, i)
        Next
        
        'read all sites into a comma delimited string
        sSites = "'" & mvStudySite(1, 0) & "'"
        For i = 1 To UBound(mvStudySite, 2)
            sSites = sSites & ",'" & mvStudySite(1, i) & "'"
        Next
        
        
        If Not bNoDBAccess Then
            'REM 17/12/03 - Check to see if the user has AllStudies/AllSites permission
            bAllStudiesAllSites = CheckAllStudiesAllSites(sUserName, sRoleCode)
            'If true set the variables to AllStudies and AllSites
            If bAllStudiesAllSites Then
                sStudies = "AllStudies"
                sSites = "AllSites"
            End If
            
            'create new studies and sites objects
            Set moStudies = New Studies
            Set moSites = New Sites
            'load all the user studies and sites
            Call moStudies.Load(mconMACRO, sStudies)
            Call moSites.Load(moDatabase.DatabaseType, mconMACRO, sSites)
        End If
    
        'loads all the studies to sites and sites to studies
        For i = 0 To UBound(mvStudySite, 2)
            'for the specified study add its linked site
            moStudies.Item(mvStudySite(0, i)).AddSite moSites.Item(mvStudySite(1, i))
            'for the specified site add the linked study
            moSites.Item(mvStudySite(1, i)).AddStudy moStudies.Item(mvStudySite(0, i))
        Next
        
        LoadStudySitePermissions = True
    Else
        
        'create new studies and sites objects
        Set moStudies = New Studies
        Set moSites = New Sites
            'load all the user studies and sites
            Call moStudies.Load(mconMACRO)
            Call moSites.Load(moDatabase.DatabaseType, mconMACRO)
        
        mvStudySite = Null
        'User has no study site permmissions under this role
        LoadStudySitePermissions = False
    End If
    
    If Not bNoDBAccess Then
        'initialise object to get lists
        Set moDataLists = New UserDataLists
    End If
    
    moDataLists.Init Me
    
Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "UserLogin.LoadStudySitePermissions"
End Function

'----------------------------------------------------------------------------------------'
Private Function CheckAllStudiesAllSites(sUserName As String, sRoleCode As String) As Boolean
'----------------------------------------------------------------------------------------'
'REM - 12/12/03
'Function to check is a user/role combo has an 'AllSites/AllStudies' permission
'returns true if a user/role has 'AllSites/AllStudies' permission
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsPermission As ADODB.Recordset
    
    On Error GoTo ErrLabel

    sSQL = "SELECT StudyCode, SiteCode" _
        & " FROM UserRole" _
        & " WHERE UserName = '" & sUserName & "'" _
        & " AND RoleCode = '" & sRoleCode & "'" _
        & " AND StudyCode = 'AllStudies'" _
        & " AND SiteCode = 'AllSites'"
    Set rsPermission = New ADODB.Recordset
    rsPermission.Open sSQL, mconMACRO, adOpenKeyset, adLockPessimistic, adCmdText

    If rsPermission.RecordCount > 0 Then
        CheckAllStudiesAllSites = True
    Else
        CheckAllStudiesAllSites = False
    End If
    
    rsPermission.Close
    Set rsPermission = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "UserLogin.CheckAllStudiesAllSites"
End Function

'----------------------------------------------------------------------------------------'
Public Function GetNewSubjectStudies() As Collection
'----------------------------------------------------------------------------------------'
'REM 30/09/02
'Returns a collection of the studies that the user has permission to access and are not closed for recruitment(3)
'----------------------------------------------------------------------------------------'
Dim oStudy As Study
Dim colStudies As Collection
    
    On Error GoTo ErrLabel
    
    Set colStudies = New Collection
    
    If CheckPermission(gsFnCreateNewSubject) Then
        
        'loop through all studies and get all that are not closed for recruitment
        For Each oStudy In moStudies.Studies
            If oStudy.StatusId <> 3 Then
                colStudies.Add oStudy
            End If
        Next
        
    End If
    
    Set GetNewSubjectStudies = colStudies
    
    Set colStudies = Nothing
    
Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "UserLogin.GetNewSubjectStudies"
End Function

'----------------------------------------------------------------------------------------'
Public Function GetNewSubjectSites(ByVal lStudyId As Long) As Collection
'----------------------------------------------------------------------------------------'
'REM 30/09/02
'Returns a collection of the sites for a specific study that the user has permission to access and the
'site status is active(0)
'----------------------------------------------------------------------------------------'
Dim oSite As Site
Dim oStudy As Study
Dim colSites As Collection
Dim colStudies As Collection
Dim sSites As String

    On Error GoTo ErrLabel

    Set colSites = New Collection
    Set colStudies = New Collection
    
    If CheckPermission(gsFnCreateNewSubject) Then
    
        'loop through all the sites user has access to
        For Each oSite In moSites.Sites
            'check if the site status is active
            If oSite.SiteStatus = 0 Then
                'get the collection of studies accociated with each site
                Set colStudies = oSite.SiteStudies
                'loop through the site's studies
                For Each oStudy In colStudies
                    'check for the passed in studyId
                    If oStudy.StudyId = lStudyId Then
                        If Not (oSite.SiteLocation = 1 And DBIsServer) Then
                            'if this is a remote site and we are on the site
                            'add site to collection
                             colSites.Add oSite
                        End If
                    End If
                 Next
            End If
        Next
        
    End If
    
    Set GetNewSubjectSites = colSites

    Set colSites = Nothing
    Set colStudies = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "UserLogin.GetNewSubjectSites"
End Function

'----------------------------------------------------------------------------------------'
Public Function GetAllStudies() As Collection
'----------------------------------------------------------------------------------------'
'REM 30/09/02
'returns a collection of all the studies that the user has permission to access
'----------------------------------------------------------------------------------------'
    
    On Error GoTo ErrLabel
      
    Set GetAllStudies = moStudies.Studies
    
Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "UserLogin.GetAllStudies"
End Function

'----------------------------------------------------------------------------------------'
Public Function GetAllSites(Optional lStudyId As Long = -1) As Collection
'----------------------------------------------------------------------------------------'
' returns a collection of all the sites that the user has permission to access
' with site code as key
'----------------------------------------------------------------------------------------'
Dim oSite As Site
Dim oStudy As Study
Dim colSites As Collection
Dim colStudies As Collection
Dim sSites As String

    On Error GoTo ErrLabel
   
    If Not moSites.Sites Is Nothing Then
         Set colSites = New Collection
        'loop through all the sites user has access to
        For Each oSite In moSites.Sites
            
            'check if the site status is active
            If oSite.SiteStatus = 0 Then
                
                If lStudyId = -1 Then 'if the StudyId = -1 then return all the sites that are active
                    colSites.Add oSite, oSite.Site
                    
                Else 'return only those for the specified studyId

                    'get the collection of studies accociated with each site
                    Set colStudies = oSite.SiteStudies
                    'loop through the site's studies
                    For Each oStudy In colStudies
                        'check for the passed in studyId
                        If oStudy.StudyId = lStudyId Then
                            'add site to collection
                            colSites.Add oSite, oSite.Site
                        End If
                    Next
                End If
            End If
        Next
    End If
    
    Set GetAllSites = colSites

    Set colSites = Nothing
    Set colStudies = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "UserLogin.GetAllSites"
End Function

'----------------------------------------------------------------------------------------'
Public Function GetOpenSubjectStudies() As Collection
'----------------------------------------------------------------------------------------'
'REM 30/09/02
'returns a collection of all the studies that the user has permission to access
'----------------------------------------------------------------------------------------'
Dim oStudy As Study
Dim colStudies As Collection
    
    On Error GoTo ErrLabel
    
    Set colStudies = New Collection
    
    'make sure the user has View Subject permission
    If CheckPermission(gsFnViewSubjectData) Then
        
        Set colStudies = moStudies.Studies

    End If
    
    Set GetOpenSubjectStudies = colStudies
    
    Set colStudies = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "UserLogin.GetOpenSubjectStudies"
End Function

'----------------------------------------------------------------------------------------'
Public Function GetOpenSubjectSites(Optional lStudyId As Long = -1) As Collection
'----------------------------------------------------------------------------------------'
'
'
'----------------------------------------------------------------------------------------'
Dim oSite As Site
Dim oStudy As Study
Dim colSites As Collection
Dim colStudies As Collection
Dim sSites As String

    On Error GoTo ErrLabel

    Set colSites = New Collection
    Set colStudies = New Collection
    
    If CheckPermission(gsFnViewSubjectData) Then
    
        'loop through all the sites user has access to
        For Each oSite In moSites.Sites
            
            'check if the site status is active
            If oSite.SiteStatus = 0 Then
                
                If lStudyId = -1 Then 'if the StudyId = -1 then return all the sites that are active
                    colSites.Add oSite
                    
                Else 'return only those for the specified studyId

                    'get the collection of studies accociated with each site
                    Set colStudies = oSite.SiteStudies
                    'loop through the site's studies
                    For Each oStudy In colStudies
                        'check for the passed in studyId
                        If oStudy.StudyId = lStudyId Then
                            'add site to collection
                            colSites.Add oSite
                        End If
                    Next
                End If
            End If
        Next
        
    End If
    
    Set GetOpenSubjectSites = colSites

    Set colSites = Nothing
    Set colStudies = Nothing
    

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "UserLogin.GetOpenSubjectSites"
End Function

'---------------------------------------------------------------------
Public Sub gLog(sUserName As String, sTaskId As String, sMessage As String, Optional conSec As ADODB.Connection = Nothing)
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
Dim oTimeZone As TimeZone
Dim SecurityCon As ADODB.Connection

    On Error GoTo ErrHandler
    
    Set oTimeZone = New TimeZone
    
    If conSec Is Nothing Then
        Set SecurityCon = mconSecurity
    Else
        Set SecurityCon = conSec
    End If
    
    'Use standard SQL datestamp
    sSQLNow = SQLStandardNow
    
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
    nTimeZone = oTimeZone.TimezoneOffset
    
    'add in location, will always be local
    sLocation = "Local"
    
    'add log
    sSQL = " INSERT INTO LoginLog " _
        & "(LogDateTime,LogNumber,TaskId,LogMessage,UserName,LogDateTime_TZ,Location,Status)" _
        & " Values (" & sSQLNow & "," & nLogNumber & ",'" & sTaskId _
        & "','" & sMessage & "','" & sUserName & "'," & nTimeZone & ",'" & sLocation & "'," & 0 & ")"

    SecurityCon.Execute sSQL
    
    Set oTimeZone = Nothing
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|" & "UserLogin.gLog"
End Sub

'---------------------------------------------------------------------
Private Function SQLStandardNow() As String
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------

    SQLStandardNow = LocalNumToStandard(IMedNow)

End Function

'---------------------------------------------------------
Public Function ReplaceQuotes(sStr As String) As String
'---------------------------------------------------------
' Reaplace all single quotes in sStr with two single quotes
' To be used for strings in SQL statements
'---------------------------------------------------------

    ReplaceQuotes = Replace(sStr, "'", "''")
    
End Function

'----------------------------------------------------------------------------------------'
Public Function GetStudiesSites(Optional sPermission As String, Optional lStudyId As Long = -1, Optional sSite As String = "", _
                                            Optional bIncludeInActiveSites As Boolean = False) As Collection
'----------------------------------------------------------------------------------------'
'REM 01/10/02
'Return collection of study site combinations for the user
'----------------------------------------------------------------------------------------'
' REVISIONS
' dph 22/03/2004 - Clean up objects
'----------------------------------------------------------------------------------------'
Dim oStudySite As StudySite
Dim colStudySites As Collection
Dim colSites As Collection
Dim colStudies As Collection
Dim oSite As Site
Dim oStudy As Study
Dim i As Integer

    On Error GoTo ErrLabel
    
    Set colStudySites = New Collection

    If CheckPermission(sPermission) Or sPermission = "" Then
        'only do this if have permission or no permission passed in- else return empty collection
        'returns all site/study combos for a specific study
        If (lStudyId <> -1) And (sSite = "") Then
            Set colSites = New Collection
            'return collection of sites for a specific study
            Set colSites = moStudies.Item(lStudyId).StudySites
            
            For Each oSite In colSites
                'only add site, study combos where the site is active
                If (oSite.SiteStatus = 0) Or bIncludeInActiveSites Then
                    Set oStudySite = New StudySite
                    With oStudySite
                        .Site = oSite.Site
                        .StudyId = lStudyId
                        .StudyName = moStudies.Item(lStudyId).StudyName
                    End With
                    colStudySites.Add oStudySite
                End If
            Next
            Set colSites = Nothing
        
        'returns all studies/site combos for a specific site
        ElseIf (sSite <> "") And (lStudyId = -1) Then
         
            'only add site, studies combos if the site is active
            If (moSites.Item(sSite).SiteStatus = 0) Or bIncludeInActiveSites Then
                Set colStudies = moSites.Item(sSite).SiteStudies
                For Each oStudy In colStudies
                    'only add sites that are active
                    If oStudy.ItemSite(sSite).SiteStatus = 0 Then
                        Set oStudySite = New StudySite
                        With oStudySite
                            .StudyId = oStudy.StudyId
                            .StudyName = oStudy.StudyName
                            .Site = moSites.Item(sSite).Site
                        End With
                        colStudySites.Add oStudySite
                    End If
                Next
                Set colStudies = Nothing
            End If
        
        'returns all study/site combos that the user has permission for
        ElseIf (lStudyId = -1) And (sSite = "") Then
            
            If Not IsNull(mvStudySite) Then
                'loop through the array that contains all the study site combo's the user has access to
                For i = 0 To UBound(mvStudySite, 2)
                    'only add if the site is active
                    If (moSites.Item(mvStudySite(1, i)).SiteStatus = 0) Or bIncludeInActiveSites Then
                        Set oStudySite = New StudySite
                        With oStudySite
                            .StudyId = mvStudySite(0, i)
                            .Site = mvStudySite(1, i)
                            .StudyName = mvStudySite(2, i)
                        End With
                        colStudySites.Add oStudySite
                    End If
                Next
            End If
            
        'returns a specific study/site combination
        ElseIf (lStudyId <> -1) And (sSite <> "") Then
            
            If (moSites.Item(sSite).SiteStatus = 0) Or bIncludeInActiveSites Then
                Set oStudySite = New StudySite
                
                With oStudySite
                    .StudyId = lStudyId
                    .StudyName = moStudies.Item(lStudyId).StudyName
                    .Site = moStudies.Item(lStudyId).ItemSite(sSite).Site
                End With
                colStudySites.Add oStudySite
            End If
            
        End If
    End If
    
    Set GetStudiesSites = colStudySites

    Set oStudySite = Nothing
    Set colStudySites = Nothing
    Set oSite = Nothing
    Set oStudy = Nothing
    
Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "UserLogin.GetStudiesSites"
End Function

'----------------------------------------------------------------------------------------'
Public Property Get DataLists() As UserDataLists
'----------------------------------------------------------------------------------------'
    
    Set DataLists = moDataLists
    
    
End Property

'----------------------------------------------------------------------------------------'
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------------------------------------------'
' REVISIONS
' TA 06/01/2003: Added UserDataLists to SerialiizedUser
'----------------------------------------------------------------------------------------'

     msDBSiteCode = PropBag.ReadProperty("DBSC")
     msSecurityCon = PropBag.ReadProperty("SC")
     msUserNameFull = PropBag.ReadProperty("UF")
     msCurrentDBConnection = DecryptString(PropBag.ReadProperty("DB"))
     msDatabaseCode = PropBag.ReadProperty("DBC")
     'nb datbase object currently only serialises the db type
     Set moDatabase = PropBag.ReadProperty("DBO")
     msUserName = PropBag.ReadProperty("US")
     msAppTitle = PropBag.ReadProperty("AP")
     Set moPermissions = PropBag.ReadProperty("PE")
     msStudySite = PropBag.ReadProperty("SS")
     Set moStudies = PropBag.ReadProperty("ST")
     Set moSites = PropBag.ReadProperty("SI")
     mbLoggedIn = PropBag.ReadProperty("LI")
     Set UserSettings = PropBag.ReadProperty("S")
     Set moDataLists = PropBag.ReadProperty("DL")
     
     'we only store the last login of the password properties
     Set moPassword = New Password
     moPassword.LastLogin = PropBag.ReadProperty("LL")
     
     LoadStudySitePermissions msUserName, moPermissions.UserRole, True

End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_Terminate()
'----------------------------------------------------------------------------------------'
' Close connections
' NCJ 27 Nov 03 - Also close down all the objects we've used
' DPH 24/03/2004 - Fix to prevent memory leak caused by cyclic references
'       References were being kept open to objects meaning they never got dropped
'       causing a gradual rise when used with web halting the web server eventually
'       Added 'Friend' Terminate events to Studies, Study, Sites, Site
'       The code follows the following steps:
'       1. The first 'pass' at terminating the objects destroys the references
'       to the 'site' objects in the collection within the 'Sites' object by
'       calling moSites.Terminate
'       2. Then call moStudies.Terminate, this has the effect of :-
'           a) Calling the Terminate method of all the 'Study' Objects in the
'               collection within 'Studies' which in turn calls the terminate
'               method of that studies' associated 'Site' objects which removes
'               their references to the 'Study' objects and destroys the collection
'           b)  Next removes all sites within the 'Study' from site collection & then
'               destroys collection
'           c)  Finally removes all 'study' objects from collection within 'Studies'
'               & destroys collection
'       Then all references are destroyed and the memory is released
'----------------------------------------------------------------------------------------'
    
    ADO_CloseConnection mconMACRO
    ADO_CloseConnection mconSecurity

    Set mconMACRO = Nothing
    Set mconSecurity = Nothing
    
    Set UserSettings = Nothing

    Call CollectionRemoveAll(mcolDatabases)
    Set mcolDatabases = Nothing
    Set moDatabase = Nothing
    Set moPasswordPolicy = Nothing
    Set moPassword = Nothing
    
    Set moPermissions = Nothing
    
    If Not (moDataLists Is Nothing) Then
        Call moDataLists.Terminate
    End If
    
    If Not (moSites Is Nothing) Then
        moSites.Terminate
    End If
    If Not (moStudies Is Nothing) Then
        moStudies.Terminate
    End If
    'Call StudiesRemoveAll(moStudies)
    'Call SitesRemoveAll(moSites)
    Set moSites = Nothing
    Set moStudies = Nothing
    Set moDataLists = Nothing
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------------------------------------------'
' REVISIONS
' TA 06/01/2003: Added UserDataLists to SerialiizedUser
'----------------------------------------------------------------------------------------'
    
    PropBag.WriteProperty "DBSC", msDBSiteCode
    PropBag.WriteProperty "SC", msSecurityCon
    PropBag.WriteProperty "UF", msUserNameFull
    PropBag.WriteProperty "DB", EncryptString(msCurrentDBConnection)
    PropBag.WriteProperty "DBC", msDatabaseCode
    PropBag.WriteProperty "DBO", moDatabase
    PropBag.WriteProperty "US", msUserName
    PropBag.WriteProperty "AP", msAppTitle
    PropBag.WriteProperty "PE", moPermissions
    PropBag.WriteProperty "SS", msStudySite
    PropBag.WriteProperty "ST", moStudies
    PropBag.WriteProperty "SI", moSites
    PropBag.WriteProperty "LI", mbLoggedIn
    PropBag.WriteProperty "S", UserSettings
    PropBag.WriteProperty "DL", moDataLists
    PropBag.WriteProperty "LL", LastLogin

    
End Sub

'----------------------------------------------------------------------------------------'
Public Function GetState(bNoDBChosen As Boolean) As String
'----------------------------------------------------------------------------------------'
' get string version of this object
'----------------------------------------------------------------------------------------'
' dph 24/03/2004 - tidy objects
'----------------------------------------------------------------------------------------'
Dim pb As PropertyBag

    If bNoDBChosen Then
        GetState = EncryptString(msUserName & "," & moPassword.Password)
    Else
        Set pb = New PropertyBag
           
        Call Class_WriteProperties(pb)
        GetState = pb.Contents
        Set pb = Nothing
    End If
    
End Function

'----------------------------------------------------------------------------------------'
Public Sub SetState(sSerialised As String)
'----------------------------------------------------------------------------------------'
' allow user object to be created from string
'----------------------------------------------------------------------------------------'
' dph 24/03/2004 - tidy objects
'----------------------------------------------------------------------------------------'
Dim pb As PropertyBag
Dim b() As Byte
    b = sSerialised
    Set pb = New PropertyBag
    pb.Contents = b
    Call Class_ReadProperties(pb)
    Set pb = Nothing
End Sub

'----------------------------------------------------------------------------------------'
Public Sub SetStateHex(sSerialised As String)
'----------------------------------------------------------------------------------------'
' allow user object to be created from string
'----------------------------------------------------------------------------------------'
    
    SetState HexToBinaryString(sSerialised)
    
End Sub

'----------------------------------------------------------------------------------------'
Public Function GetStateHex(bNoDBChosen As Boolean) As String
'----------------------------------------------------------------------------------------'
' return serialised user as hex string
'----------------------------------------------------------------------------------------'

    GetStateHex = BinaryStringToHex(GetState(bNoDBChosen))
    
End Function


'----------------------------------------------------------------------------------------'
Public Property Get LastLogin() As Double
'----------------------------------------------------------------------------------------'
'make the user's last login date available publically
'----------------------------------------------------------------------------------------'

    LastLogin = moPassword.LastLogin
    
End Property


'---------------------------------------------------------------------
Public Function ShowOCDiscMenu() As Boolean
'---------------------------------------------------------------------
'return if access is allowed to the OCDISCREPANCIES View menu item
'---------------------------------------------------------------------
Dim rs As ADODB.Recordset

    ShowOCDiscMenu = False
    
    If CheckPermission(gsFnCreateDiscrepancy) And moDatabase.DatabaseType = Oracle80 Then
            'allowed to raise discrepancies and this is an oracle database
        Set rs = New ADODB.Recordset
        rs.Open "SELECT COUNT(*) FROM ALL_VIEWS WHERE VIEW_NAME = 'OCDISCREPANCIES'", mconMACRO
        If rs.Fields(0).Value = 1 Then
            'the OCDISCREPANCIES view exists - allow access to it
            ShowOCDiscMenu = True
        End If
        rs.Close
        Set rs = Nothing
    End If

End Function


Public Function GetSubjectLabel(lStudyId As Long, sSite As String, lSubjectId As Long) As String
'---------------------------------------------------------------------
' NCJ 7 Aug 2000 Retrieve Subject label given StudyId, Site and SubjectID
' Returns empty string if no label exists
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler

    If mconMACRO Is Nothing Then
        Set mconMACRO = New Connection
        mconMACRO.Open msCurrentDBConnection
        mconMACRO.CursorLocation = adUseClient
    End If
    
    sSQL = "SELECT LocalIdentifier1 FROM TrialSubject " _
        & " WHERE ClinicalTrialId = " & lStudyId _
        & " AND TrialSite = '" & sSite & "'" _
        & " AND PersonId = " & lSubjectId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, mconMACRO, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        GetSubjectLabel = ""
    Else
        GetSubjectLabel = ConvertFromNull(rsTemp!LocalIdentifier1, vbString)
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description & "|MACROUser.GetSubjectLabel"
End Function
