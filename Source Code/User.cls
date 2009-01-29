VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       ClsUser.cls
'   Author:     Andrew Newbigging, November 1997
'   Purpose:    Used to check user security details and return collection of user.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'
'   1    Andrewn    19/11/97  15:52    Created Clsuser.cls
'   2    Andrewn    27/11/97  15:11    Checked in $/macro/Security
'   3   Andrew Newbigging   8/10/98     SPR 427
'       New public proporty DatabaseName added for use in setting the
'       reports to the database selected by the user (using an ODBC DSN of the
'       same name as the DatabaseName)
'       Mo Morris           5/1/99      SPR 663
'UCase() function calls added to Public function Login for the purpose of making the logon
'case-insensitive. This facility was not required on Username because the SQL statement that
'performs a check between the entered Username and those on file works in a case-insensitive manner.
'       Andrew newbigging   26/2/99
'   Added new property indicating the type of database the user has logged on to
'       Mo Morris           6/8/99
'       ReadUserDatabase changed, all references to vUserName changed to vUserCode
'       ReadFunctionAccess Temporarily commented out (until Role/RoleFunction/UserRole work is ready)
'       CheckFunctionAccess Temporarily changed to always return true (until Role/RoleFunction/UserRole work is ready)
'   4   Paul Norris         03/09/99    Added SecurityDatabase() Property Get
'   5   Paul Norris         13/09/99    Table name User changed to MacroUser in Security.mdb
'                                       Updated code to conform to v1.0 of coding standards
'   6   PN  13/09/99    Added SecurityDatabasePath() property get and amended OpenSecurityDatabase()
'                       to use it
'   7   WillC 8/10/99   Added HasPasswordExpired() to test if the login password is valid
'   8   Willc 11/10/99  Added IsPasswordDefault() to test if the login is using the defaultlogin
'                       if so prompt the user to change the password.
'   9 NCJ 15 Oct 99 - Rewrote Case statement in HasPasswordExpired
'                   Also use today's date rather than LastLogin
'   NCJ 25 Oct 99   Changed UserName to UserCode where necessary
'                   NB Global property UserName here is actually UserCode
'   Mo Morris   27/10/99
'               Private variables  msServerName and msNameOfDatabase created
'               Property Get/Let procedures for ServerName and NameOfDatabase
'   Mo Morris   29/10/99 - DAO to ADO conversion started
'               OpenSecurityDatabase removed, now handled by modADODBConnection.InitializeSecurityADODBConnection
'               2/11/99  - DAO to ADO conversion completed
'   NCJ 3 Dec 99    Use single quotes in SQL strings
'   NCJ 10 Dec 99 - Enabled the CheckFunctionAccess function
'   NCJ 14 Dec 99 - Set up user role and functions after setting up database,
'                   and look for user's role in specific database
'   NCJ 17 Dec 99 - Store FirstLogin as double
'   NCJ 20 Dec 99 - Made GetRole a Public Function
'   ATN 14/1/2000 - Added databaseuser and databasepassword and password
'   NCJ 24/1/00      - Added SecurityDatabasePath Let, and DefaultSecurityDatabasePath
'   NCJ 25/1/00   - Maintain "local" copy of Security DB path for current session
'   Mo Morris   2/5/00,   Default value of 'AutoImport' added to UserName Property Get
'   TA 18/08/2000 UserHasAccessToModule function moved here from frmLogin to help WWW development
'   Mo 18/7/2001  Changes stemming from field Password in table MacroUser being changed to
'               UserPassword (stems from the swith to Jet 4.0)
'   DPH 25/10/2001 - Fix to use new password entered rather than old one on change
'   TA 06/02/2002: VTRACK Changes for default security path
'   Mo Morris 11/6/2002 Changes to UserHasAccessToModule for new Query Module
'   TA 14/08/2002: Changes to SetDatabasePath to crete a new MACRO database when no dbs exist
'   ASH 12/9/2002 - Registry keys replaced with calls to new Settings file
'   ASH 12/9/2002 Changed DatabaseName to NameOfDatabase
'----------------------------------------------------------------------------------------'

Option Explicit
Option Base 0
Option Compare Binary

Private msRole As String
Private msRoleCode As String    'SDM SR1222 22/11/99 Authentication
Private msUserFullName As String
Private msUserName As String
Private msDatabasepath As String
Private msDatabaseName As String
Private msDatabaseUser As String
Private msDatabasePassword As String
Private msPassword As String
' NCJ 25/1/00
Private msSecurityDatabasePath As String

'   ATN 26/2/99
'   Added property for the database type
Private mnDatabaseType As MACRODatabaseType

'added by Mo Morris 27/10/99
Private msServerName As String
Private msNameOfDatabase As String

Private mcolUserDatabases As MACRODatabases

Private mcolUserFunctions As Collection

'Public Enum LoginResult
'    Success = 0
'    AccountDisabled = 1
'    Fail = 2
'End Enum

Public gnCollectioncount As Integer

'   ATN 13/12/99
'   Not used by ADO.
'Private mdbSecurity As Database

'SDM 14/01/00   This should not be exposed.
''----------------------------------------------------------------------------------------'
'Public Property Get UserFunctions() As Collection
''----------------------------------------------------------------------------------------'
'' PN 13/09/99 Changed public in line properties to conform with coding standard v1.0
''----------------------------------------------------------------------------------------'
'
'    Set UserFunctions = mcolUserFunctions
'
'End Property

'----------------------------------------------------------------------------------------'
Public Property Get DatabaseType() As MACRODatabaseType
'----------------------------------------------------------------------------------------'

    DatabaseType = mnDatabaseType

End Property

'----------------------------------------------------------------------------------------'
Public Property Let DatabaseType(nDatabaseType As MACRODatabaseType)
'----------------------------------------------------------------------------------------'

    mnDatabaseType = nDatabaseType

End Property

'----------------------------------------------------------------------------------------'
Public Property Get UserName() As String
'----------------------------------------------------------------------------------------'
' NB This is actually the User Code
'----------------------------------------------------------------------------------------'
'Changed Mo Morris 2/5/00   When Trial Offices /AI Auto import or /TR are being used, there
'needs to be a UserName available when DataItemResponse tables are changed/updated. These will
'now get a UserName of "AutoImport". (Note: it was thought that /TR would not get used much
'and did not need to be reflected in the choice of UserName)
'----------------------------------------------------------------------------------------'

    If RTrim(msUserName) = "" Then
        UserName = "AutoImport"
    Else
        UserName = msUserName
    End If

End Property

'----------------------------------------------------------------------------------------'
Public Property Let UserName(sUsername As String)
'----------------------------------------------------------------------------------------'
' NB This is actually the User Code
'----------------------------------------------------------------------------------------'

    msUserName = sUsername

End Property

'----------------------------------------------------------------------------------------'
Public Property Get Password() As String
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'

    Password = msPassword

End Property

'----------------------------------------------------------------------------------------'
Public Property Let Password(sPassword As String)
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'

    msPassword = sPassword

End Property

'----------------------------------------------------------------------------------------'
Public Property Get DefaultSecurityDatabasePath() As String
'----------------------------------------------------------------------------------------'
' Get MACRO's default security path (i.e. the one set on installation)
'----------------------------------------------------------------------------------------'

'TA 06/02/2002: VTRACK Changes for default security path
#If VTRACK = 1 Then
    DefaultSecurityDatabasePath = App.Path & "\databases\VTRACKSecurity.mdb"
#Else
    DefaultSecurityDatabasePath = App.Path & "\databases\security.mdb"
#End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get SecurityDatabasePath() As String
'----------------------------------------------------------------------------------------'
' PN 15/09/99 - security db path is obtained from this new property get
' Get the security database path as stored in the Registry
' NCJ 25/1/00 - Store as local value which won't be changed for the duration of this session
' (NB The stored registry value might be changed by System Management, but we need
' to be able to reinitialise the Security DB connection to what we started with)
' ASH 12/9/2002 - Registry keys replaced with calls to new Settings file
'----------------------------------------------------------------------------------------'
'   ATN 14/1/2000
'   Read security path from the registry
'  NCJ 24/1/00 Use GetSecurityPathKey

    ' Check if local value already set up
    'If msSecurityDatabasePath = "" Then
        ' Read value from Registry
        ' SDM 26/01/00 SR2794
        'msSecurityDatabasePath = GetFromRegistry(GetMacroRegistryKey, "SecurityPath")
        ' sSecurityPath = GetFromRegistry(GetSecurityPathKey, "SecurityPath")
        'If Len(msSecurityDatabasePath) = 0 Then
            ' No registry value - use "default"
            'msSecurityDatabasePath = DefaultSecurityDatabasePath
        'End If
        'msSecurityDatabasePath = DefaultSecurityDatabasePath
        'msSecurityDatabasePath = GetMACROSetting("SecurityPath", "")
     'End If
        ' ASH 11/9/2002 Get security database path from the Settings file
        'msSecurityDatabasePath = GetMACROSetting("SecurityPath", "")
     'End If
    ' Return value already set up
    'SecurityDatabasePath = msSecurityDatabasePath
    
    
    ' ASH 11/9/2002 Get security database path from the Settings file
    ' using IMEDSettings component.Returns the stored security database path
    ' location or the default.
    If msSecurityDatabasePath = "" Then
        msSecurityDatabasePath = GetMACROSetting("SecurityPath", DefaultSecurityDatabasePath)
    End If
    
    ' Return value already set up
    SecurityDatabasePath = msSecurityDatabasePath
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Let SecurityDatabasePath(sNewSecurityPath As String)
'----------------------------------------------------------------------------------------'
' NCJ 24/1/00 - Allow setting of default path (e.g. from System Management)
' Assume sNewSecurityPath is valid
' Note this does NOT change the local value
' ASH 12/9/2002 - Registry keys replaced with calls to new Settings file
'----------------------------------------------------------------------------------------'
Dim sRegPath As String

    'SDM 26/01/00 SR2794
    'sRegPath = GetMacroRegistryKey
    'sRegPath = GetSecurityPathKey
    'SetKeyValue sRegPath, "SecurityPath", sNewSecurityPath, REG_SZ
    
    'ASH 11/9/2002 Now using new IMEDSettings component to add to settings file
    Call SetMACROSetting("SecurityPath", sNewSecurityPath)
    

End Property

'SDM 26/01/00 SR2794
''----------------------------------------------------------------------------------------'
'Private Function GetSecurityPathKey() As String
''----------------------------------------------------------------------------------------'
'' NCJ 24/1/00
'' Get the name of the Registry key for the Security database path
'' Specify the text for the key elements (ProductName, CompanyName)
'' in case Project Properties are inadvertently changed
''----------------------------------------------------------------------------------------'
'Dim sVersion As String
'Dim sAppName As String
'Dim sCompany As String
'
''   ATN 14/1/2000
''   Read security path from the registry
'
'    sVersion = App.Major
''    sAppName = App.ProductName
''    sCompany = App.CompanyName
'    sAppName = "MACRO"
'    sCompany = "InferMed Limited"
'    GetSecurityPathKey = "Software\" & sCompany & "\" & sAppName & "\" & sVersion
'
'End Function

'----------------------------------------------------------------------------------------'
Public Property Get DatabasePath() As String
'----------------------------------------------------------------------------------------'

    DatabasePath = msDatabasepath

End Property

'----------------------------------------------------------------------------------------'
Public Property Let DatabasePath(sDatabasePath As String)
'----------------------------------------------------------------------------------------'

    msDatabasepath = sDatabasePath

End Property

'----------------------------------------------------------------------------------------'
Public Property Get DatabaseName() As String
'----------------------------------------------------------------------------------------'

    DatabaseName = msDatabaseName

End Property

'----------------------------------------------------------------------------------------'
Public Property Let DatabaseName(sDatabaseName As String)
'----------------------------------------------------------------------------------------'

    msDatabaseName = sDatabaseName

End Property

'   ATN 13/12/99
'   Property not used anywhere.
''----------------------------------------------------------------------------------------'
'Public Property Get SecurityDatabase() As Database
''----------------------------------------------------------------------------------------'
'' PN change - 03/09/99
'' expose the security database as a property
''----------------------------------------------------------------------------------------'
'
''    Set SecurityDatabase = mdbSecurity
'
'End Property

'----------------------------------------------------------------------------------------'
Public Function login(ByVal sUserCode As String, ByVal sPassword As String, _
                      Optional ByVal vCheckPassword As Boolean, Optional bCheckDetailsOnly As Boolean = True) As Long
'----------------------------------------------------------------------------------------'
' Try to log in the user with user code sUserCode and password sPassword
' If user doesn't exist or if password is wrong return "fail" result
' NCJ 17 Dec 99 - Store FirstLogin for new user as double
'TA 15/05/2000 - now returns as failed if password expired and not changed
'----------------------------------------------------------------------------------------'
Dim rsPasswords As ADODB.Recordset
Dim rsUser As ADODB.Recordset
Dim nExpiryPeriod
Dim nNewExpiryPeriod
Dim sSQL As String
Dim sSQL1 As String
Dim sFirstLogin As String
Dim dtfirstlogin As Date
' Dim dtLastLogin As Date
    
    On Error GoTo ErrHandler

    'changed by Mo Morris 2/10/99
    'now handled by modADODBConnection.InitializeSecurityADODBConnection
    'Call OpenSecurityDatabase
    
    ' PN 13/09/99 field name User changed to MacroUser
    ' NCJ 25 Oct 99 Change UserName to UserCode
    ' sSQL = "SELECT * FROM MacroUser WHERE UserName = '" & sUserCode & "'"
    'Mo Morris 20/9/01 Db Audit (UserCode to UserName)
    sSQL = "SELECT * FROM MacroUser WHERE UserName = '" & sUserCode & "'"
    Set rsUser = New ADODB.Recordset
    rsUser.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    If rsUser.RecordCount < 1 Then
    ' User does not exist
        login = LoginResult.Failed
        Exit Function
    Else
    
        'Call to ReadUserDatabase uses UserCode instead of UserName
        Call ReadUserDatabase(sUserCode)
        Call UpdateLastLogin
        login = LoginResult.Success
    End If
    
'    sSQL1 = "SELECT * FROM MacroPassword "
'    Set rsPasswords = New ADODB.Recordset
'    rsPasswords.Open sSQL1, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
'    nExpiryPeriod = rsPasswords!ExpiryPeriod
'    rsPasswords.Close
'    Set rsPasswords = Nothing
'
'    nNewExpiryPeriod = 90
'    If nExpiryPeriod = 0 Then
'        sSQL = "UPDATE MacroPassword SET " _
'            & " ExpiryPeriod = " & nNewExpiryPeriod
'        SecurityADODBConnection.Execute sSQL
'    End If
'
'    'dtFirstLogin = rsUser!FirstLogin
'    If IsNull(rsUser!FirstLogin) Then       ' New database/ new user?
'
'
'        ' NCJ 20 Oct 99 - Make sure we only update path for default "Access" DB
'        'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode)
'        sSQL = "UPDATE Databases SET DatabaseLocation = '" & App.Path _
'            & "\databases\macro.mdb" & "'" _
'            & " WHERE DatabaseCode = 'Access'"
'        SecurityADODBConnection.Execute sSQL
'
''        dtFirstLogin = Format(Now, "dd/mm/yyyy hh:mm:ss")
'        dtfirstlogin = Now
'        ' Update the MacroUser table for this user
'
'    ' WillC 4/2/00 Changed to cope with local settings
'    sFirstLogin = ConvertLocalNumToStandard(CStr(CDbl(dtfirstlogin)))
'
'        ''Mo Morris 20/9/01 Db Audit (UserCode to UserName)
'         sSQL = "UPDATE MacroUser SET " _
'                & " FirstLogin = " & sFirstLogin _
'                & " WHERE UserName = '" & sUserCode & "'"
'         SecurityADODBConnection.Execute sSQL
'    Else
'        dtfirstlogin = CDate(rsUser!FirstLogin)
'    End If
'    ' Don't need to check LastLogin (NCJ 25/10/99)
''    If IsNull(rsUser!LastLogin) Then
''        dtLastLogin = Format(Now, "dd/mm/yyyy hh:mm:ss")
''    Else
''        dtLastLogin = rsUser!LastLogin
''    End If
'
'    ' Moved HasPasswordExpired to further down - NCJ 25/10/99
'    'UCase calls added by Mo Morris 5/1/99 SPR 663
'    With rsUser
'        If .RecordCount > 0 Then
'            If rsUser!Enabled = 0 Then
'                login = LoginResult.AccountDisabled
'            'changed Mo 18/7/2001
'            'ElseIf UCase(rsUser!Password) = UCase(sPassword) Or vCheckPassword = False Then
'            ElseIf UCase(rsUser!UserPassword) = UCase(sPassword) Or vCheckPassword = False Then
'                ' Set global UserName property (misnamed!) to User Code
                UserName = sUserCode
                Password = sPassword
'                'ZA 21/05/2005, changed UserName to UserNameFull
'                UserFullName = .Fields("UserNameFull").Value
'                login = LoginResult.Success
'                Call IsPasswordDefault(sUserCode, sPassword)
'                ' Password Expiry moved here from before password check - NCJ 25/10/99
'                If HasPasswordExpired(sUserCode, dtfirstlogin) Then
'                    'password has expired and
'                    ' PN 20/09/99
'                    ' this may be a timeout login so do not do all initialization again if it is
'                    If Not bCheckDetailsOnly Then
'                        'Changed by Mo Morris 6/8/99
'                        'Call to ReadUserDatabase uses UserCode instead of UserName
'                        Call ReadUserDatabase(sUserCode)
'                        Call UpdateLastLogin
'                    End If
'                Else
'                    login = LoginResult.Fail
'                End If
'            Else
'                login = LoginResult.Fail
'
'            End If
'
'        Else
'            login = LoginResult.Fail
'
'        End If
'
'        .Close
'
'    End With
'
     'Set up this user's functions
     'AFTER the user has selected a database
     Call FetchUserFunctions
'
'    Set rsUser = Nothing
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Login", "clsUser")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

''----------------------------------------------------------------------------------------'
'Private Sub OpenSecurityDatabase()
''----------------------------------------------------------------------------------------'
'
'    ' PN 15/09/99 - read the security db path from the new property get SecurityDatabasePath()
'    ' Establish connection to database
'    Set mdbSecurity = DBEngine.Workspaces(0).OpenDatabase(SecurityDatabasePath)
'
'End Sub

'----------------------------------------------------------------------------------------'
Private Sub UpdateLastLogin()
'----------------------------------------------------------------------------------------'
' Record the last login of the current user
'----------------------------------------------------------------------------------------'

Dim sSQL As String
        
    On Error GoTo ErrHandler


    ' PN 13/09/99 table name User changed to MacroUser
    ' NCJ 25/10/99 Change UserName column name to UserCode (global UserName still OK)
    ' WillC 4/2/00 Changed LastLogin to SQLStandardNow to cope with locale separators
    ' in doubles (datetimes)
    'Mo Morris 20/9/01 Db Audit (UserCode to UserName)
    sSQL = "UPDATE MacroUser " _
        & " SET LastLogin = " & SQLStandardNow _
        & " WHERE UserName = '" & UserName & "'"
        
              ' & " SET LastLogin = '" & Format(Now, "dd/mm/yyyy hh:mm:ss")
              
    SecurityADODBConnection.Execute sSQL
        
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpdateLastLogin", "clsUser")
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
Private Sub ReadUserDatabase(vUserCode As String)
'----------------------------------------------------------------------------------------'
'Changed by Mo Morris 6/8/99
'All references to vUserName changed to vUserCode
'----------------------------------------------------------------------------------------'

Dim sSQL As String
Dim rsUserDatabase As ADODB.Recordset
        
    On Error GoTo ErrHandler

    'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode, UserCode to UserName)
    sSQL = "SELECT DatabaseCode FROM UserDatabase WHERE UserName = '" & vUserCode & "'"
    Set rsUserDatabase = New ADODB.Recordset
    rsUserDatabase.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Set mcolUserDatabases = Nothing
    Set mcolUserDatabases = New MACRODatabases
    
    mcolUserDatabases.Clear
    
    With rsUserDatabase
        Do While Not .EOF
            mcolUserDatabases.AddUserDatabase rsUserDatabase!DataBaseCode
            .MoveNext
        Loop
        
        .Close
    End With
    
    Set rsUserDatabase = Nothing
        
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ReadUserDatabase", "clsUser")
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
Public Sub SetDatabasePath(vDatabase As String)
'----------------------------------------------------------------------------------------'
'changed by Mo Morris 27/10/99
'ServerName and NameOfDatabase added to SQLQuery of Databases together with
'their Property Let calls
' NCJ 14 Dec 99 - Do GetRole and FetchUserFunctions here (after storing database details)
' DPH 10/04/2002 - Added SecureHTMLLocation setting
'----------------------------------------------------------------------------------------'
Dim sPath As String
Dim rsDatabase As ADODB.Recordset
Dim sSQL As String
Dim sString As String
    
    On Error GoTo ErrHandler

    'changed by Mo Morris 27/10/99
    '   ATN 14/1/2000   Added databaseuser and databasepassword
    'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode)
    ' DPH 10/04/2002 - Added SecureHTMLLocation
    sSQL = "SELECT DatabaseLocation,DatabaseType,ServerName,NameOfDatabase,DatabaseUser,DatabasePassword, " _
            & "HTMLLocation, SecureHTMLLocation FROM Databases " _
                & " WHERE DatabaseCode = '" & vDatabase & "'"
    Set rsDatabase = New ADODB.Recordset
    rsDatabase.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    'TA 15/08/2002: all null comparison changed to use removenull
    DatabasePath = RemoveNull(rsDatabase!DatabaseLocation)
    
    DatabaseType = rsDatabase!DatabaseType
    DatabaseName = vDatabase
    DatabaseUser = RemoveNull(rsDatabase!DatabaseUser)
    DatabasePassword = RemoveNull(rsDatabase!DatabasePassword)
    ServerName = RemoveNull(rsDatabase!ServerName)
    NameOfDatabase = RemoveNull(rsDatabase!NameOfDatabase)
    
    'SDM 26/01/00 SR2794
    If Not IsNull(rsDatabase!HTMLLocation) Then
        gsHTML_FORMS_LOCATION = rsDatabase!HTMLLocation
    Else
        gsHTML_FORMS_LOCATION = gsAppPath & "html\"
    End If
    ' DPH 10/04/2002 - New Secure HTML folder (corrected 22/04/2002)
    If RemoveNull(rsDatabase!SecureHTMLLocation) = "" Then
        ' default to HTML location
        gsSECURE_HTML_LOCATION = gsHTML_FORMS_LOCATION
    Else
        gsSECURE_HTML_LOCATION = rsDatabase!SecureHTMLLocation
    End If

    rsDatabase.Close
    Set rsDatabase = Nothing
    
    'TA 16/08/2002: check a database exists
    'if it's the first time MACRO is run a blank MSDE db needs to be created
    If DatabaseType = sqlserver And ServerName = "NEED TO CREATE" Then
        If Not CreateFirstDatabase Then
            'this returns true if there is already a database or one is successfully created
            ExitMACRO
            MACROEnd
        End If
    End If
    
    'connect to the database and test that it's a Macro Database
    Dim ADODBConnection As ADODB.Connection
    Set ADODBConnection = New ADODB.Connection
    Select Case DatabaseType
    Case MACRODatabaseType.Access
            'JET OLE DB NATIVE PROVIDER
            ADODBConnection.Open Connection_String(CONNECTION_MSJET_OLEDB_40, DatabasePath, , , DatabasePassword)
    Case MACRODatabaseType.sqlserver
            'SQL SERVER OLE DB NATIVE PROVIDER
            'ASH 12/9/2002 Changed DatabaseName to NameOfDatabase
            ADODBConnection.Open Connection_String(CONNECTION_SQLOLEDB, ServerName, NameOfDatabase, _
                DatabaseUser, DatabasePassword)
'   ATN 21/12/99
'   Added connection to Oracle
    Case MACRODatabaseType.Oracle80
            'Oracle OLE DB NATIVE PROVIDER
            'Mo Morris 20/8/01, Connection string call uses NameOfDatabase instead of ServerName
            ADODBConnection.Open Connection_String(CONNECTION_MSDAORA, NameOfDatabase, , _
                DatabaseUser, DatabasePassword)
    End Select
    
    Call GetRole    'SDM 11/11/99
    Call FetchUserFunctions
    
     ' NCJ 17 Dec 99
     ' Moved here from GetRole 20/12/99
    #If WebRDE <> -1 Then
    Select Case App.Title
            Case "MACRO_DM", "MACRO_SM"
                frmMenu.sbrMenu.Panels.Item("UserKey").Text = "Current User: " & gUser.UserFullName & " (" & gUser.UserName & ")"
                frmMenu.sbrMenu.Panels.Item("RoleKey").Text = "Role: " & gUser.Role
     End Select
    #End If
    
'    sSQL = "Insert Into DataType (DataTypeId, DataTypeName) values (0,'Text')"
'    ADODBConnection.Execute sSQL
    

Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "SetDatabasePath", Err.Description

'    Select Case Err.Description
'    Case "Couldn't find output table 'DataType'."
'         MsgBox "This is not Macro database, please choose a valid" + vbCr _
'              & " Macro database.", vbInformation, "MACRO"
'         frmLogin.Show vbModal
'         Exit Sub
'    End Select
'
'On Error GoTo 0
 
End Sub

'----------------------------------------------------------------------------------------'
Public Property Get UserDatabases() As MACRODatabases
'----------------------------------------------------------------------------------------'

    Set UserDatabases = mcolUserDatabases

End Property

'----------------------------------------------------------------------------------------'
Public Function CheckFunctionAccess(ByVal sFunction As String) As Boolean
'----------------------------------------------------------------------------------------'
' Check the current user's access rights against the given function string
' (See modUserFunctions for the available function strings)
' The mcolUserFunctions collection is set up during login
'----------------------------------------------------------------------------------------'
Dim vFunction As Variant

Dim sReturnFunction As String

    ' Default to false
    CheckFunctionAccess = False
    
    'SDM 14/01/00   I have altered FetchUserFunction to put the function code
    '               in as the key as well. This allows us to see if the function
    '               is in the collection without having to loop through it.
    '               I felt the need...the need for speed.
    On Error Resume Next
    sReturnFunction = mcolUserFunctions(sFunction)
    If Err.Number = 0 Then  'No error, its in there.
        CheckFunctionAccess = True
    End If
    On Error GoTo 0
    
'    ' See if the requested function is in the collection
'    ' of this user's functions
'    For Each vFunction In mcolUserFunctions
'        If CStr(vFunction) = sFunction Then
'            ' We've found the one we want
'            CheckFunctionAccess = True
'            Exit Function
'        End If
'    Next
End Function

'------------------------------------------------------------------------------------'
Public Function UserHasAccessToModule() As Boolean
'------------------------------------------------------------------------------------'
' Returns TRUE if user's role allows access to this application
' Assume current user's function collection already set up
' NCJ 9 Dec 99
' TA 18/08/2000 moved here from frmLogin to help WWW development
'------------------------------------------------------------------------------------'
Dim sAppTitle As String

    On Error GoTo ErrHandler
    
    ' This gets filled with app title if they don't have access
    sAppTitle = ""
    
    Select Case App.Title
    
    Case "MACRO_SM"
        If CheckFunctionAccess(gsFnSystemManagement) = False Then
            sAppTitle = "MACRO System Management"
        End If
        
    Case "MACRO_EX"
        If CheckFunctionAccess(gsFnExchange) = False Then
            sAppTitle = "MACRO Exchange"
        End If
        
   Case "MACRO_DM"
        If LCase$(Command) = "review" Then
            If CheckFunctionAccess(gsFnDataReview) = False Then
                sAppTitle = "MACRO Data Review"
            End If
        Else
            If CheckFunctionAccess(gsFnDataEntry) = False Then
                sAppTitle = "MACRO Data Entry"
            End If
        End If
        
    Case "MACRO_SD"
        If LCase$(Command) = "library" Then
            If CheckFunctionAccess(gsFnLibraryManagement) = False Then
                sAppTitle = "MACRO Library Management"
            End If
        Else
            If CheckFunctionAccess(gsFnStudyDefinition) = False Then
                sAppTitle = "MACRO Study Definition"
            End If
        End If
    'ta 18/04/2001: check for create data views
    Case "MacroCreateDataViews"
        If Not CheckFunctionAccess(gsFnCreateDataViews) Then
            sAppTitle = "MACRO Create Data Views"
        End If
    'Mo 11/6/2002
    Case "MACRO_QM"
        If CheckFunctionAccess(gsFnQueryModule) = False Then
            sAppTitle = "MACRO Query Module"
        End If
    End Select

    If sAppTitle > "" Then  ' They don't have access
        MsgBox "You do not have permission to access " & sAppTitle, _
                vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, sAppTitle
        UserHasAccessToModule = False
    Else
        UserHasAccessToModule = True
        ' Let frmMenu enable its menu options etc.
        ' Need definition of CheckUserRights in each frmMenu
        Call frmMenu.CheckUserRights
    
    End If
    
    Exit Function
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "UserHasAccessToModule")
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
Private Function HasPasswordExpired(sUserCode As String, dtfirstlogin As Date) As Boolean
'----------------------------------------------------------------------------------------'
'Find out if the last login is greater than the first login in days as set in
'the MacroPassword table if greater than expiry period then disable the user,
'if within 7,6,5 etc days flash a message saying how many days to change.
' NCJ 15 Oct 99 - Use "now" instead of dtLastLogin; removed dtLastLogin as parameter
' TA 15/05/2000 - now returns False if password expired and not changed
' NCJ 19/10/00 - Tidied up messages
'----------------------------------------------------------------------------------------'

Dim sSQL As String
Dim nExpiryPeriod As Integer
Dim nValidDaysLeft As Integer
Dim nDifference As Integer
Dim rsPasswords As ADODB.Recordset
Dim sMsg As String
Dim bSuccess As Boolean
Dim sNewPassword As String

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM MacroPassword"
    Set rsPasswords = New ADODB.Recordset
    rsPasswords.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    nExpiryPeriod = rsPasswords!ExpiryPeriod
    Set rsPasswords = Nothing
    
    ' Find no. of days from first login with this password to today's date
    nDifference = DateDiff("d", dtfirstlogin, Now)
    
    ' Find no. of valid days left with this password
    ' (may be negative if password has expired)
    nValidDaysLeft = nExpiryPeriod - nDifference
           
    'fail only occurs if user fails to change password AND password has expired
    bSuccess = True
           
    ' DPH 25/10/2001 - Reset Password
    sNewPassword = ""
    
    Select Case nValidDaysLeft
        Case 2 To 7
            sMsg = "You have " & nValidDaysLeft & " days before your password expires." _
                    & vbCr & "Do you wish to change it now?"
            If MsgBox(sMsg, vbQuestion + vbYesNo, gsDIALOG_TITLE) = vbYes Then
                Call frmChangePassword.Display(sUserCode, sNewPassword)
            End If
        Case 1
            sMsg = "Your password will expire after today - do you wish to change it now?"
            If MsgBox(sMsg, vbQuestion + vbYesNo, gsDIALOG_TITLE) = vbYes Then
                Call frmChangePassword.Display(sUserCode, sNewPassword)
            End If
        Case Is < 1
            ' Password has expired - they must change it now (NCJ 15 Oct 99)
            MsgBox "Your password has expired, and it must be changed now.", vbInformation, "MACRO"
            bSuccess = frmChangePassword.Display(sUserCode, sNewPassword)
    End Select
    
    ' DPH 25/10/2001 - storing any password change in class member
    If sNewPassword <> "" Then
        Me.Password = sNewPassword
    End If
    
    HasPasswordExpired = bSuccess
     
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "HasPasswordExpired", "clsUser")
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
Private Sub DisablePassword(sUser As String)
'----------------------------------------------------------------------------------------'
' Set the enabled flag to 0 to disable the user
' sUser is User Code
'----------------------------------------------------------------------------------------'

Dim sSQL As String
    
    On Error GoTo ErrHandler

    'Mo Morris 20/9/01 Db Audit (UserCode to UserName)
    sSQL = "UPDATE MacroUser " _
    & " SET Enabled = 0" _
    & " WHERE UserName = '" & sUser & "'"
            
    SecurityADODBConnection.Execute sSQL
        
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DisablePassword", "clsUser")
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
Private Sub IsPasswordDefault(sUser As String, sPassword As String)
'----------------------------------------------------------------------------------------'
' check to see if the user who has logged in is the default User
' UserCode is rde and the password is macrotm
'----------------------------------------------------------------------------------------'
'REVISIONS
' 25/10/2001 - Added new password setting
'----------------------------------------------------------------------------------------'
Dim sMsg As String
Dim sNewPassword As String

    On Error GoTo ErrHandler
    
    ' DPH 25/10/2001 - reset new password
    sNewPassword = ""
    
    If LCase(sUser) = "rde" And LCase(sPassword) = "macrotm" Then
        sMsg = " You have logged on using the default username and password." & vbCr _
             & " Would you like to change the default password now?"
        Select Case MsgBox(sMsg, vbQuestion + vbYesNo, gsDIALOG_TITLE)
                Case vbYes
                    Call frmChangePassword.Display(sUser, sNewPassword)
                    ' DPH 25/10/2001 - storing any password change in class member
                    If sNewPassword <> "" Then
                        Me.Password = sNewPassword
                    End If

                Case vbNo
        End Select
    End If
        
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "IsPasswordDefault", "clsUser")
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
Public Property Get ServerName() As String
'----------------------------------------------------------------------------------------'

    ServerName = msServerName

End Property

'----------------------------------------------------------------------------------------'
Public Property Let ServerName(ByVal sServerName As String)
'----------------------------------------------------------------------------------------'

    msServerName = sServerName

End Property

'----------------------------------------------------------------------------------------'
Public Property Get NameOfDatabase() As String
'----------------------------------------------------------------------------------------'

    NameOfDatabase = msNameOfDatabase
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Let NameOfDatabase(ByVal sNameOfDatabase As String)
'----------------------------------------------------------------------------------------'

    msNameOfDatabase = sNameOfDatabase

End Property

Public Property Get UserFullName() As String
    UserFullName = msUserFullName
End Property

Private Property Let UserFullName(ByVal sNewValue As String)
    msUserFullName = sNewValue
End Property

Public Property Get Role() As String
    Role = msRole
End Property

Private Property Let Role(ByVal sNewValue As String)
    msRole = sNewValue
End Property

'------------------------------------------------------
Public Property Get RoleCode() As String
'------------------------------------------------------
'SDM SR1222 22/11/99 Authentication
'------------------------------------------------------
    
    RoleCode = msRoleCode

End Property

'------------------------------------------------------
Private Property Let RoleCode(ByVal sNewValue As String)
'------------------------------------------------------

    msRoleCode = sNewValue

End Property

'------------------------------------------------------
Public Property Get DatabasePassword() As String
'------------------------------------------------------

    
    DatabasePassword = msDatabasePassword

End Property

'------------------------------------------------------
Public Property Let DatabasePassword(ByVal sNewValue As String)
'------------------------------------------------------

    msDatabasePassword = sNewValue

End Property

'------------------------------------------------------
Public Property Get DatabaseUser() As String
'------------------------------------------------------

    
    DatabaseUser = msDatabaseUser

End Property

'------------------------------------------------------
Public Property Let DatabaseUser(ByVal sNewValue As String)
'------------------------------------------------------

    msDatabaseUser = sNewValue

End Property

'------------------------------------------------------
Public Function GetRole() As String
'------------------------------------------------------
' Get user's role
' NB The DatabaseName must be already set up
' NCJ 20/12/99 - Made Public Function (to be called from frmTempLogin)
' Returns current user's RoleCode
'------------------------------------------------------
Dim sSQL As String
Dim rsRole As ADODB.Recordset
    
    On Error GoTo ErrHandler

    ' NCJ 14 Dec 99 - Added DatabaseDescription to WHERE clause
    ' (a user's role applies to a specific database)
    'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode, UserCode to UserName)
    sSQL = "SELECT Role.RoleDescription, Role.RoleCode " & _
           "FROM Role, UserRole " & _
           "WHERE Role.RoleCode = UserRole.RoleCode " & _
           "AND UserRole.UserName = '" & UserName & "'" & _
           "AND UserRole.DatabaseCode = '" & DatabaseName & "'"
           
    Set rsRole = New ADODB.Recordset
    rsRole.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not rsRole.EOF Then
        rsRole.MoveFirst
        Role = RemoveNull(rsRole.Fields("RoleDescription").Value)
        RoleCode = RemoveNull(rsRole.Fields("RoleCode").Value)  'SDM SR1222 22/11/99 Authentication
    Else
        Role = ""
        RoleCode = ""
    End If
    Set rsRole = Nothing
    
    ' NCJ 20/12/99
   GetRole = RoleCode
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "GetRole", "clsUser")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

'-----------------------------------------
Private Function FetchUserFunctions()
'-----------------------------------------
' Create the collection of User Functions
' NB this can only be done AFTER a database has been selected
' and the RoleCode has been set up
'-----------------------------------------
Dim rsUserFunctions As ADODB.Recordset
Dim sSQL As String
Dim sFnCode As String

    sSQL = "SELECT * FROM RoleFunction " _
            & " WHERE RoleCode = '" & gUser.RoleCode & "'"
        
    Set rsUserFunctions = New ADODB.Recordset
    rsUserFunctions.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    
    Set mcolUserFunctions = New Collection

    'SDM 14/01/00   I have added the function code as the key as well for faster searching.
    With rsUserFunctions
        Do Until .EOF = True
            sFnCode = rsUserFunctions!FunctionCode
'            mcolUserFunctions.Add CStr(rsUserFunctions!FunctionCode), rsUserFunctions!FunctionCode
            mcolUserFunctions.Add sFnCode, sFnCode
               .MoveNext
'            Debug.Print "Added function -" & sFnCode & "- len = " & Len(sFnCode)
        Loop
    End With

    rsUserFunctions.Close
    Set rsUserFunctions = Nothing
        
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "FetchUserFunctions", "clsUser")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function

'-----------------------------------------
Private Sub Class_Initialize()
'-----------------------------------------
' NCJ 25/1/00 - Initialise Security DB Path
'-----------------------------------------

    msSecurityDatabasePath = ""
    
End Sub
