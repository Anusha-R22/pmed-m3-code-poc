Attribute VB_Name = "modDatabaseUpgrade"
'----------------------------------------------------------------------------------------'
'   File:       modDatabaseUpgrade.bas
'   Copyright:  InferMed Ltd. 2003. All Rights Reserved
'   Author:     Richard Meinesz, June 2003
'   Purpose:    Utilitiy routines for upgradeing a database from pervious versions of MACRO
'----------------------------------------------------------------------------------------'
' Revisions:
'----------------------------------------------------------------------------------------'

Option Explicit

'these are required when we use late binding on the MACROAccess22ToMSDE22.Xfer class
Private Enum eFailReason
    sfrSuccess
    sfrNotAccess22
    sfrNotLocal
    sfrWrongSaPwd
    sfrFailPhysicalCreate
    sfrFailMACROCreate
    sfrFailMTS
    sfrFailCreateUser
    sfrFailUpdateSecurity
    sfrFailMarkAccessUpdated
End Enum

'---------------------------------------------------------------------------------
Public Sub UpgradeToLatestMACRODatabase(sSecCon As String, sDatabaseCode As String, sDBVersion As String, _
                                        ByRef sMessage As String)
'---------------------------------------------------------------------------------
'REM 10/06/03
'Upgrade a MACRO database to the latest MACRO 3.0 database
'---------------------------------------------------------------------------------
Dim sSQL As String
Dim rsDatabase As ADODB.Recordset
Dim oSecConnection As ADODB.Connection
Dim oDBConnection As ADODB.Connection
Dim nDatabaseType As Integer
Dim sConnection As String
Dim sAccessCon As String
Dim sDatabasePath As String
Dim sDatabasePswd As String
Dim sMSDECon As String
Dim sErrMsg As String
Dim sUpgradePath As String
Dim oUpgradeDB20To22 As UpMACRO20To22
    
    On Error GoTo Errlabel
    
    HourglassOn
    
    'create security database connection
    Set oSecConnection = CreateDatabaseConnection(sMessage, sSecCon)
    
    If sMessage <> "" Then Exit Sub

    sSQL = "SELECT * FROM Databases WHERE DatabaseCode = '" & sDatabaseCode & "'"
    Set rsDatabase = New ADODB.Recordset
    rsDatabase.Open sSQL, oSecConnection, adOpenKeyset, adLockPessimistic, adCmdText
    
    If rsDatabase.RecordCount <> 0 Then
        nDatabaseType = rsDatabase!DatabaseType
        
        Set oUpgradeDB20To22 = New UpMACRO20To22
        
        'create database connection string
        Select Case nDatabaseType
        Case MACRODatabaseType.Access
            'ACCESS
            sDatabasePswd = DecryptString(rsDatabase!DatabasePassword)
            sDatabasePath = rsDatabase!DatabaseLocation
            sAccessCon = Connection_String(CONNECTION_MSJET_OLEDB_40, sDatabasePath, , , sDatabasePswd)
            
            'First upgrade the Access database to latest version of MACRO 2.2
            If oUpgradeDB20To22.Init(sErrMsg, sAccessCon, Access, sDatabasePath, sDatabasePswd) Then
                Call oUpgradeDB20To22.UpgradeDataDatabase20To22(nDatabaseType)
                
                'Convert it to MACRO 2.2 MSDE database
                If UpgradeAccess22ToMSDE22(sSecCon, sAccessCon, sDatabaseCode, sMSDECon) Then
                
                    If UpgradeMACRO22ToMACRO30(sSecCon, sMSDECon, sDatabaseCode, sqlserver, "2.0.17", sErrMsg) Then 'use 2.0.17 as a fake DB version to ensure that the correct upgrade scripts are run
                        Call DialogInformation("The Access database has been succefully upgraded to a MACRO 3.0 SQL Server/MSDE database")
                    Else
                        Call DialogError(sErrMsg)
                    End If
                End If
            Else
                Call DialogError("An error occured while trying to connect to the Access database. " & vbCrLf & sErrMsg & vbCrLf & "Upgrade aborted!")
            End If
            HourglassOff
            
        Case MACRODatabaseType.sqlserver
            
            sDatabasePswd = RemoveNull(rsDatabase!DatabasePassword)
            If sDatabasePswd <> "" Then
               sDatabasePswd = DecryptString(sDatabasePswd)
            End If

            'SQL SERVER OLE DB NATIVE PROVIDER
            sConnection = Connection_String(CONNECTION_SQLOLEDB, rsDatabase!ServerName, rsDatabase!NameOfDatabase, _
                          DecryptString(rsDatabase!DatabaseUser), sDatabasePswd)

            'Upgrade to latest version of MACRO 2.2
            If oUpgradeDB20To22.Init(sErrMsg, sConnection, sqlserver) Then
                Call oUpgradeDB20To22.UpgradeDataDatabase20To22(nDatabaseType)
                'Upgrade to 3.0 from latest version of MACRO 2.2 (Routine can also upgrade different versions of 3.0)
                If UpgradeMACRO22ToMACRO30(sSecCon, sConnection, sDatabaseCode, sqlserver, sDBVersion, sErrMsg) Then
                    Call DialogInformation("The MACRO database has been successfully upgraded to MACRO 3.0.")
                Else
                    DialogError sErrMsg
                End If
            Else
                Call DialogError("An error occured while trying to connect to the database. " & vbCrLf & sErrMsg & vbCrLf & "Upgrade aborted!")
            End If
            HourglassOff
            
        Case MACRODatabaseType.Oracle80
            
            sDatabasePswd = RemoveNull(rsDatabase!DatabasePassword)
            If sDatabasePswd <> "" Then
               sDatabasePswd = DecryptString(sDatabasePswd)
            End If
            
            'Oracle OLE DB NATIVE PROVIDER
            sConnection = Connection_String(CONNECTION_MSDAORA, rsDatabase!NameOfDatabase, , _
                          DecryptString(rsDatabase!DatabaseUser), DecryptString(rsDatabase!DatabasePassword))
            
            'Upgrade to latest version of MACRO 2.2
            If oUpgradeDB20To22.Init(sErrMsg, sConnection, Oracle80) Then
                Call oUpgradeDB20To22.UpgradeDataDatabase20To22(nDatabaseType)
                'Upgrade to 3.0 from latest version of MACRO 2.2 (Routine can also upgrade different versions of 3.0)
                If UpgradeMACRO22ToMACRO30(sSecCon, sConnection, sDatabaseCode, Oracle80, sDBVersion, sErrMsg) Then
                    Call DialogInformation("The MACRO database has been successfully upgraded to MACRO 3.0.")
                Else
                    DialogError sErrMsg
                End If
            Else
                Call DialogError("An error occured while trying to connect to the database. " & vbCrLf & sErrMsg & vbCrLf & "Upgrade aborted!")
            End If
            HourglassOff
        End Select
    
    End If

Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|modUpgradeDatabases.UpgradeToLatestMACRODatabase"
End Sub

'--------------------------------------------------------------------------------------------------
Private Function UpgradeMACRO22ToMACRO30(sSecCon As String, sMACROCon As String, sDatabaseCode As String, _
                                    eDatabaseType As MACRODatabaseType, sDatabaseVersion As String, _
                                    Optional sErrMsg As String = "") As Boolean
'--------------------------------------------------------------------------------------------------
'REM 16/06/03
'Routine to upgrade an MSDE, SQL Serever or Oracle MACRO 2.2 final database to a MACRO 3.0 database
'--------------------------------------------------------------------------------------------------
Dim sScriptPrefix As String
Dim lBuildSubVersion As Long
Dim sBuildSubVersion As String
Dim nBuildSubVersion As Integer
Dim nBuildVersion As Integer
Dim oMACROCon As ADODB.Connection
Dim sUpgradePath As String
Dim vDBVersion As Variant
    
    On Error GoTo Errlabel

    vDBVersion = Split(sDatabaseVersion, ".")
    nBuildVersion = vDBVersion(0)
    nBuildSubVersion = vDBVersion(2)
    
    If nBuildVersion <> 3 Then 'pre-MACRO 3.0 databse and will not contain a UserRole table so create it
        'Add UserRole Table and user roles to the upgraded MACRO database
        If MACRODBUserRoleTable(sSecCon, sMACROCon, sDatabaseCode, eDatabaseType, sErrMsg) Then
            'If had to create UserRole table then need to change SubVersion number to 17 so scripts will run
            nBuildSubVersion = "17"
        Else
            GoTo Errlabel
        End If
    End If
    
    'Then upgrade to latest version of MACRO 3.0
    Set oMACROCon = CreateDatabaseConnection(sErrMsg, sMACROCon)
    If sErrMsg <> "" Then GoTo Errlabel
    
    sUpgradePath = App.Path & "\Database Scripts\Upgrade Database\"
    
    Select Case eDatabaseType
    Case MACRODatabaseType.sqlserver
        sScriptPrefix = "MSSQL"
    Case MACRODatabaseType.Oracle80
        sScriptPrefix = "ORA"
    End Select
    
    For lBuildSubVersion = nBuildSubVersion To (CURRENT_SUBVERSION - 1)
    
        sBuildSubVersion = CStr(lBuildSubVersion + 1)
        ExecuteMultiLineSQL oMACROCon, _
                            StringFromFile(sUpgradePath & sScriptPrefix & "_30_" & CStr(lBuildSubVersion) _
                            & "To" & sBuildSubVersion & ".sql")
    Next
    
    UpgradeMACRO22ToMACRO30 = True
    
Exit Function
Errlabel:
    If sErrMsg = "" Then
        sErrMsg = "Error while running MACRO 3.0 upgrade script (" & sScriptPrefix & "_30_" & CStr(lBuildSubVersion) & "TO" & sBuildSubVersion & ".sql)" & ". Error Description: " & Err.Description & " Error Number: " & Err.Number
    End If
    UpgradeMACRO22ToMACRO30 = False
End Function

'--------------------------------------------------------------------------------------------------
Private Function CreateDatabaseConnection(ByRef sErrorMsg As String, sCon As String) As ADODB.Connection
'--------------------------------------------------------------------------------------------------
'REM 11/06/03
'Create a database connection
'--------------------------------------------------------------------------------------------------
Dim oConnection As ADODB.Connection

    On Error GoTo Errlabel

    Set oConnection = New ADODB.Connection
    oConnection.Open sCon
    oConnection.CursorLocation = adUseClient
    
    Set CreateDatabaseConnection = oConnection
   
Exit Function
Errlabel:
    sErrorMsg = "Database connection error: " & Err.Description & ": Error no. " & Err.Number
End Function

'---------------------------------------------------------------------
Private Function MACRODBUserRoleTable(sSecCon As String, sMACROCon As String, sDatabaseCode As String, _
                                 eMACRODatabaseType As MACRODatabaseType, sErrMsg As String) As Boolean
'---------------------------------------------------------------------
'REM 16/06/03
' Add UserRole table to the MACRO database from the Security Db
' The Upgraded security database will contain a table called MACRO22UserRole
' this contains all user roles from MACRO 2.2 and here these are converted into MACRO 3.0 user roles
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsMACROControl As ADODB.Recordset
Dim rsUserRole As ADODB.Recordset
'Dim sUserNameDT As String
'Dim sRoleCodeDT As String
'Dim sStudySiteDT As String
'Dim sInstallDT As String
Dim oSecCon As ADODB.Connection
Dim oMACROCon As ADODB.Connection
Dim lErrNo As Long
Dim sUsername As String
Dim sUserRole As String

    On Error GoTo Errlabel
    
    'create the user role table in the upgraded database
    Call CreateUserRoleTable(sMACROCon, eMACRODatabaseType, sErrMsg)
    If sErrMsg <> "" Then GoTo Errlabel

    'create database and security database connections
    Set oSecCon = CreateDatabaseConnection(sErrMsg, sSecCon)
    Set oMACROCon = CreateDatabaseConnection(sErrMsg, sMACROCon)
    If sErrMsg <> "" Then GoTo Errlabel
    
    On Error Resume Next
    'get the user roles from the old security UserRole table (now in the upgraded sec db in the MACRO22UserRole table)
    'that are for the current MACRO database being upgraded
    sSQL = "SELECT DISTINCT UserName, RoleCode FROM MACRO22UserRole" _
        & " WHERE DatabaseCode = '" & sDatabaseCode & "'"
    Set rsUserRole = New ADODB.Recordset
    rsUserRole.Open sSQL, oSecCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    lErrNo = Err.Number
    Err.Clear
    
    On Error GoTo Errlabel
    'If above errors then there is no MACRO22UserRole tabel in the security database, which means it is a security database that was not upgraded
    'from an earlier verion on MACRO but created in MACRO 3.0
    'and the database being upgraded was registered in MACRO 3.0 but was an earlier version of MACRO
    If lErrNo = 0 Then
        'add the roles to the new MACRO UserRole table and add the database code and user name to the User Dtabase table
        Do While Not rsUserRole.EOF
            
            sUsername = rsUserRole!UserName
            sSQL = "INSERT INTO UserRole VALUES ('" & sUsername & "','" & rsUserRole!RoleCode & "', 'AllStudies', 'AllSites', 0)"
            oMACROCon.Execute sSQL
            
            Call UpdateUserDatabases(sUsername, sDatabaseCode, oSecCon)
            
            rsUserRole.MoveNext
            
        Loop
        
        rsUserRole.Close
        Set rsUserRole = Nothing
        
    Else 'MACRO22UserRole Table does not exist in this sec database, so just add current logged on user
        
        sUsername = goUser.UserName
        sUserRole = goUser.UserRole
        sSQL = "INSERT INTO UserRole VALUES ('" & sUsername & "','" & sUserRole & "','AllStudies','AllSites',0)"
        oMACROCon.Execute sSQL
        
        'Check if users needs to be added to UserDatabase table
        Call UpdateUserDatabases(sUsername, sDatabaseCode, oSecCon)
        
    End If
    
    MACRODBUserRoleTable = True
    
    oMACROCon.Close
    Set oMACROCon = Nothing
    
    oSecCon.Close
    Set oSecCon = Nothing
    
Exit Function
Errlabel:
    If sErrMsg = "" Then
        sErrMsg = "Error while inserting MACRO 2.2 User Role table. Error Description: " & Err.Description & " Error Number: " & Err.Number
    End If
    MACRODBUserRoleTable = False
End Function

'--------------------------------------------------------------------------------
Private Sub CreateUserRoleTable(sMACROCon As String, eMACRODatabaseType As MACRODatabaseType, sErrMsg As String)
'--------------------------------------------------------------------------------
'REM 05/08/03
'Create the User Role table in a separat routine and close the connection.
'Do this due to errors in inserting into a newly created table in Oracle (Breaking the connection commits the new table)
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim sUserNameDT As String
Dim sRoleCodeDT As String
Dim sStudySiteDT As String
Dim sInstallDT As String
Dim oMACROCon As ADODB.Connection
Dim rsUR As ADODB.Recordset
Dim lErrNo As Long
    
    On Error GoTo Errlabel

    Select Case eMACRODatabaseType
    Case MACRODatabaseType.sqlserver
        sUserNameDT = "VARCHAR(20)"
        sRoleCodeDT = "VARCHAR(15)"
        sStudySiteDT = "VARCHAR(50)"
        sInstallDT = "SMALLINT"
    Case MACRODatabaseType.Oracle80
        sUserNameDT = "VARCHAR2(20)"
        sRoleCodeDT = "VARCHAR2(15)"
        sStudySiteDT = "VARCHAR2(50)"
        sInstallDT = "NUMBER(6)"
    End Select
    
    Set oMACROCon = CreateDatabaseConnection(sErrMsg, sMACROCon)
    If sErrMsg <> "" Then GoTo Errlabel
    
    'Check to see if the DB has a UserRole table
    On Error Resume Next
    sSQL = "SELECT * FROM UserRole"
    Set rsUR = New ADODB.Recordset
    rsUR.Open sSQL, oMACROCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    lErrNo = Err.Number
    Err.Clear
    rsUR.Close
    
    On Error GoTo Errlabel
    
    If lErrNo <> 0 Then 'if there was an error then it doesn't exist so create
    
        '***Add New UserRole Table***
        sSQL = "CREATE Table UserRole(UserName " & sUserNameDT & "," _
            & " RoleCode " & sRoleCodeDT & "," _
            & " StudyCode " & sStudySiteDT & "," _
            & " SiteCode " & sStudySiteDT & "," _
            & " TypeOfInstallation " & sInstallDT & "," _
            & " CONSTRAINT PKUserRole PRIMARY KEY (UserName,RoleCode,StudyCode,SiteCode,TypeOfInstallation))"
        oMACROCon.Execute sSQL
        
        '*** Add New UserRole table to the MACROTable ***
        sSQL = "INSERT INTO MACROTable VALUES ('UserRole','',0,0,0)"
        oMACROCon.Execute sSQL
        
    End If
    
    Set oMACROCon = Nothing
    
Exit Sub
Errlabel:
If sErrMsg = "" Then
    sErrMsg = "Error while creating User Role Table. Error Description: " & Err.Description & " Error Number: " & Err.Number
End If
End Sub

'--------------------------------------------------------------------------------
Private Sub UpdateUserDatabases(sUsername As String, sDatabaseCode As String, oSecCon As ADODB.Connection)
'--------------------------------------------------------------------------------
'REM 23/10/02
'Checks to see if the User/Database combination selected exists in the UserDatabase table
' if not create a new row
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsUserDB As ADODB.Recordset

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM UserDatabase" _
        & " WHERE UserName = '" & sUsername & "'" _
        & " AND DatabaseCode = '" & sDatabaseCode & "'"
    Set rsUserDB = New ADODB.Recordset
    rsUserDB.Open sSQL, oSecCon, adOpenKeyset, adLockOptimistic, adCmdText
    
    If rsUserDB.RecordCount = 0 Then
        sSQL = "INSERT INTO UserDatabase " _
            & " VALUES ('" & sUsername & "','" & sDatabaseCode & "')"
            oSecCon.Execute sSQL, , adCmdText
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmNewUserRole.UpdateUserDatabases"
End Sub

''---------------------------------------------------------------------
'Public Function MSDEUpgrade() As Boolean
''---------------------------------------------------------------------
'' Upgrade a 2.2 access db to 2.2 MSDE
'' Retunrs false if unsuccessful
''---------------------------------------------------------------------
''late binding used so that changes to MACROAccess22ToMSDE22.Xfer do not require rerefernces
''switch to early binding when coding for intellisense and compilation errors
'Dim oMSDE As MACROAccess22ToMSDE22.Xfer     'MACROAccess22ToMSDE22.Xfer
'Dim sLocalMachineName As String
'Dim sSAPassword As String
'Dim enFailReason As Long   'MACROAccess22ToMSDE22.eFailReason
'Dim sDBCode As String
'Dim sSecCon As String
'Dim sMsg As String
'Dim sAccessCon As String
'Dim sLoginMsg As String
'Dim sMSDEUserId As String
'
'    'default to false
'    MSDEUpgrade = False
'
'    sDBCode = goUser.Database.NameOfDatabase
'    sSecCon = Connection_String(CONNECTION_MSJET_OLEDB_40, SecurityDatabasePath, , , gsSecurityDatabasePassword)
'
''check whether it is access
'    If goUser.Database.DatabaseType <> MACRODatabaseType.Access Then
'        'if not access then no work needed
'        MSDEUpgrade = True
''EXIT FUNCTION HERE
'        Exit Function
'    End If
'
''late binding used so that changes to MACROAccess22ToMSDE22.Xfer do not require rerefernces
''switch to early binding when coding for intellisense and compilation errors
'    Set oMSDE = CreateObject("MACROAccess22ToMSDE22.Xfer")  'New MACROAccess22ToMSDE22.Xfer
'
''is it the latest verion of MACRO 2.2 and Access>
'    If Not oMSDE.IsDBAccess22(gsADOConnectString) Then
''EXIT FUNCTION HERE
'        Exit Function
'    End If
'
''is db on local machine
'    If Not oMSDE.IsDBLocal(gsADOConnectString) Then
'        DialogInformation "This database is not local and must be upgraded manually"
''EXIT FUNCTION HERE
'        Exit Function
'    End If
'
'    sLocalMachineName = oMSDE.GetLocalHostName
'    sSAPassword = "" ' this may be change by CheckSAPassword
'
'    'ask them if they want to proceed
'    sMsg = "This is a version 2.2 Access database it must be converted to MSDE before continuing"
'    sMsg = sMsg & vbCrLf & "Do you wish to attempt this automatically now?"
'    If DialogQuestion(sMsg, "MACRO Upgrade") = vbNo Then
'        Exit Function
'    End If
'
''check sa password and whether MSDE is running locally
'    If Not oMSDE.CheckSAPassword(sLocalMachineName, sSAPassword, True, gsADOConnectString) Then
'        DialogInformation "sa password is incorrect or an instance of MSDE is not running on this machine"
''EXIT FUNCTION
'        Exit Function
'    End If
'
'    enFailReason = oMSDE.AutoXfer(sDBCode, sSecCon, sMSDEUserId, sSAPassword, , sLocalMachineName)
'    If enFailReason = sfrSuccess Then  '
'        MSDEUpgrade = True
'    Else
'        DialogInformation oMSDE.GetFailReasonText(enFailReason)
'    End If
'
'    'update guser object
'    'goUser.Database.DatabaseType = sqlserver
'    'goUser.Database.NameOfDatabase = sDBCode
'    'goUser.Database.ServerName = sLocalMachineName
'    'goUser.Database.DatabaseUser = sDBCode
'
'    Call InitializeMacroADODBConnection(False)
''    Set goUser = New MACROUser
''    goUser.login Connection_String(CONNECTION_MSJET_OLEDB_40, gUser.SecurityDatabasePath, , , gsSecurityDatabasePassword), _
''            gUser.UserName, _
''            gUser.Password, _
''            "", "", sLoginMsg, , gUser.DatabaseName, _
''            gUser.RoleCode
'
'
'End Function

'---------------------------------------------------------------------
Public Function UpgradeAccess22ToMSDE22(sSecCon As String, sAccessDBCon As String, sDBCode As String, ByRef sMSDECon As String) As Boolean
'---------------------------------------------------------------------
'REM 11/06/03
'Upgrade a MACRO 2.2 Access database to an MSDE 2.2 database
'---------------------------------------------------------------------
Dim oMSDE As MACROAccess22ToMSDE22.Xfer
Dim sLocalMachineName As String
'Dim sMSDEUserId As String
'Dim sMSDEPswd As String
Dim enFailReason As Long   'MACROAccess22ToMSDE22.eFailReason
Dim sMSG As String
Dim sAccessCon As String
Dim sLoginMsg As String
    
    'default to false
    UpgradeAccess22ToMSDE22 = False
    
    Set oMSDE = New MACROAccess22ToMSDE22.Xfer
        
    'is it the latest verion of MACRO 2.2 and Access>
    If Not oMSDE.IsDBAccess22(sAccessDBCon) Then
        DialogInformation "This is not a MACRO Access database"
        UpgradeAccess22ToMSDE22 = False
        'EXIT FUNCTION HERE
        Exit Function
    End If
    
    'is db on local machine
    If Not oMSDE.IsDBLocal(sAccessDBCon) Then
        DialogInformation "This database is not local and must be upgraded manually"
        UpgradeAccess22ToMSDE22 = False
        'EXIT FUNCTION HERE
        Exit Function
    End If
    
    HourglassSuspend
'    'will create new MSDE MACRO
    sMSDECon = frmConnectionString.Display(False, True, False, , , True)
    
    HourglassResume
''****'NB - need to ask user for MSDE username and password!!!!!!!!!!!*******
'    sLocalMachineName = oMSDE.GetLocalHostName
'    sMSDEPswd = "macrotm"
'    sMSDEUserId = "sa"
    
    'ask them if they want to proceed
'    sMsg = "This is a version 2.2 Access database it must be converted to MSDE before continuing"
'    sMsg = sMsg & vbCrLf & "Do you wish to attempt this automatically now?"
'    If DialogQuestion(sMsg, "MACRO Upgrade") = vbNo Then
'        Exit Function
'    End If
    
'    'check sa password and whether MSDE is running locally
'    If Not oMSDE.CheckSAPassword(sLocalMachineName, sMSDEPswd, True, sAccessDBCon) Then
'        DialogInformation "The MSDE password is incorrect or an instance of MSDE is not running on this machine"
'        'EXIT FUNCTION
'        Exit Function
'    End If
    
    If sMSDECon <> "" Then
        'Do upgrade
        enFailReason = oMSDE.AutoXfer(sDBCode, sSecCon, sAccessDBCon, sMSDECon)
        If enFailReason = sfrSuccess Then  '
            UpgradeAccess22ToMSDE22 = True
        Else
            DialogInformation oMSDE.GetFailReasonText(enFailReason)
            UpgradeAccess22ToMSDE22 = False
        End If
    
        'WHY IS THIS FALSE?????????
        Call InitializeMacroADODBConnection(False)
    Else 'user must have cancelled out of frmConnectionString so no database created
        UpgradeAccess22ToMSDE22 = False
    End If

End Function

'--------------------------------------------------------------------------------
Public Sub ImportSecurityData(sSecurityPath As String, sNewSecurityDBCon As String)
'--------------------------------------------------------------------------------
'REM 09/06/03
'Routine to import security data from an old Access pre-MACRO 3.0 security database to
' a new MACRO 3.0 security database (MSDE/SQL Server or Oracle)
'--------------------------------------------------------------------------------
Dim oNewSecCon As ADODB.Connection
Dim oOldSecCon As ADODB.Connection
Dim sSQL As String
Dim rs22Database As ADODB.Recordset
Dim rs30Database As ADODB.Recordset
Dim oSecUpgrade As UpSecurity20To22
Dim nDatabaseType As Integer
    
    On Error GoTo Errlabel
    
    HourglassOn
    
    'create connection to new security database
    Set oNewSecCon = New ADODB.Connection
    oNewSecCon.Open sNewSecurityDBCon
    oNewSecCon.CursorLocation = adUseClient
    
    Select Case Connection_Property(CONNECTION_PROVIDER, sNewSecurityDBCon)
    Case CONNECTION_MSDAORA, CONNECTION_ORAOLEDB_ORACLE
        nDatabaseType = MACRODatabaseType.Oracle80
    Case Else
        nDatabaseType = MACRODatabaseType.sqlserver
    End Select

    'create connection to old Access security database
    Set oOldSecCon = New ADODB.Connection
    oOldSecCon.Open Connection_String(CONNECTION_MSJET_OLEDB_40, sSecurityPath, , , gsSecurityDatabasePassword)
    
    'Upgrade Secuirity database to MACRO 2.2 first
    Set oSecUpgrade = New UpSecurity20To22
    Call oSecUpgrade.Init(oOldSecCon, Access)
    oSecUpgrade.UpgradeSecurityDatabaseTo22
    
    Set oSecUpgrade = Nothing
    
    'IMPORT SECURITY DATA

'************** MACROUser info ***********************
Dim sPswd As String
Dim sPswdAndDate  As String
Dim sPswdCreated As String
Dim sHashedPswdAndDate As String
Dim sUsername As String
Dim sLastLogin As String
Dim sFirstLogin As String
Dim sUserNameFull As String
    
    sSQL = "SELECT * FROM MACROUser"
    Set rs22Database = New ADODB.Recordset
    rs22Database.Open sSQL, oOldSecCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    If rs22Database.RecordCount <> 0 Then
        
        Do While Not rs22Database.EOF
            
            sUsername = rs22Database!UserName
            
            'REM 31/07/03 - check user name in Uppercase in Oracle as it is case sensitive
            Select Case Connection_Property(CONNECTION_PROVIDER, sNewSecurityDBCon)
            Case CONNECTION_MSDAORA, CONNECTION_ORAOLEDB_ORACLE
                sSQL = "SELECT * FROM MACROUser WHERE upper(UserName) = upper('" & sUsername & "')"
            Case Else
                sSQL = "SELECT * FROM MACROUser WHERE UserName = '" & sUsername & "'"
            End Select
    
            Set rs30Database = New ADODB.Recordset
            rs30Database.Open sSQL, oNewSecCon, adOpenKeyset, adLockPessimistic, adCmdText
            
            'check to see if user name already exists
            If rs30Database.RecordCount = 0 Then
                    
                sPswd = rs22Database!UserPassword
                
                sPswdCreated = LocalNumToStandard(IMedNow)
            
                sPswdAndDate = sPswd & sPswdCreated
            
                'hash the new password and create date
                sHashedPswdAndDate = HashHexEncodeString(sPswdAndDate)
                
                sFirstLogin = RemoveNull(rs22Database!FirstLogin)
                If sFirstLogin = "" Then sFirstLogin = LocalNumToStandard(IMedNow)
                sLastLogin = RemoveNull(rs22Database!LastLogin)
                If sLastLogin = "" Then sLastLogin = sFirstLogin
                
                sUserNameFull = RemoveNull(rs22Database!UserNameFull)
                If sUserNameFull = "" Then sUserNameFull = sUsername
                
                sSQL = "INSERT INTO MACROUser (USERNAME,USERNAMEFULL,USERPASSWORD,ENABLED,LASTLOGIN," _
                    & " FIRSTLOGIN,DEFAULTUSERROLECODE,FAILEDATTEMPTS,PASSWORDCREATED,SYSADMIN) " _
                    & " VALUES ('" & sUsername & "','" & sUserNameFull & "','" & sHashedPswdAndDate _
                    & "'," & 1 & "," & sLastLogin & "," & sFirstLogin & ",'" & "" & "'," & 0 & "," _
                    & sPswdCreated & "," & 0 & ")"
                oNewSecCon.Execute sSQL
                
                'Also insert the password into the password history table
                sSQL = "INSERT INTO PasswordHistory (USERNAME,HISTORYNUMBER,PASSWORDCREATED,USERPASSWORD) " _
                    & " VALUES ('" & rs22Database!UserName & "'," & 1 & "," & sPswdCreated & ",'" & sHashedPswdAndDate & "')"
                oNewSecCon.Execute sSQL
                
            End If
            rs22Database.MoveNext
        Loop
    End If
    

'************** Databases info *********************
Dim sDatabaseCode As String
Dim lDatabaseType As String
Dim sReportsLocation As String
Dim sEncryptedDatabasePswd As String
Dim sEncryptedDatabaseUser As String

    sSQL = "SELECT * FROM Databases "
    Set rs22Database = New ADODB.Recordset
    rs22Database.Open sSQL, oOldSecCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    If rs22Database.RecordCount <> 0 Then
    
        Do While Not rs22Database.EOF
            
            sDatabaseCode = rs22Database!DatabaseCode
            
            'REM 31/07/03 - check Database code in Uppercase in Oracle as it is case sensitive
            Select Case Connection_Property(CONNECTION_PROVIDER, sNewSecurityDBCon)
            Case CONNECTION_MSDAORA, CONNECTION_ORAOLEDB_ORACLE
                sSQL = "SELECT * FROM Databases WHERE upper(DatabaseCode) = upper('" & sDatabaseCode & "')"
            Case Else
                sSQL = "SELECT * FROM Databases WHERE DatabaseCode = '" & sDatabaseCode & "'"
            End Select
            
            Set rs30Database = New ADODB.Recordset
            rs30Database.Open sSQL, oNewSecCon, adOpenKeyset, adLockPessimistic, adCmdText
            
            'Check to see if database code already exists in new database
            If rs30Database.RecordCount = 0 Then
            
                lDatabaseType = rs22Database!DatabaseType
            
                sEncryptedDatabasePswd = EncryptString(rs22Database!DatabasePassword)
                If lDatabaseType = MACRODatabaseType.Access Then
                    sEncryptedDatabaseUser = ""
                Else
                    sEncryptedDatabaseUser = EncryptString(rs22Database!DatabaseUser)
                End If
            
                sReportsLocation = ""
                
                sSQL = "INSERT INTO Databases (DATABASECODE,HTMLLOCATION,DATABASELOCATION,DATABASETYPE,SERVERNAME," _
                    & " NAMEOFDATABASE,DATABASEUSER,DATABASEPASSWORD,SECUREHTMLLOCATION,REPORTSLOCATION) " _
                    & " VALUES ('" & sDatabaseCode & "','" & rs22Database!HTMLLocation & "','" _
                    & rs22Database!DatabaseLocation & "'," & lDatabaseType & ",'" & rs22Database!ServerName & "','" _
                    & rs22Database!NameOfDatabase & "','" & sEncryptedDatabaseUser & "','" & sEncryptedDatabasePswd & "','" _
                    & rs22Database!SecureHTMLLocation & "','" & sReportsLocation & "')"
                oNewSecCon.Execute sSQL
            
            End If
            rs22Database.MoveNext
        Loop
        
    End If

'*************************** User Database ***********************

    sSQL = "SELECT * FROM UserDatabase "
    Set rs22Database = New ADODB.Recordset
    rs22Database.Open sSQL, oOldSecCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    If rs22Database.RecordCount <> 0 Then
    
        Do While Not rs22Database.EOF
        
            sUsername = rs22Database!UserName
            sDatabaseCode = rs22Database!DatabaseCode
            
            sSQL = "SELECT * FROM UserDatabase "
            
            'REM 31/07/03 - check in Uppercase in Oracle as it is case sensitive
            Select Case Connection_Property(CONNECTION_PROVIDER, sNewSecurityDBCon)
            Case CONNECTION_MSDAORA, CONNECTION_ORAOLEDB_ORACLE
                sSQL = sSQL & " WHERE upper(DatabaseCode) = upper('" & sDatabaseCode & "')" _
                            & " AND upper(UserName) = upper('" & sUsername & "')"
            Case Else
                sSQL = sSQL & " WHERE DatabaseCode = '" & sDatabaseCode & "'" _
                            & " AND UserName = '" & sUsername & "'"
            End Select
                
            Set rs30Database = New ADODB.Recordset
            rs30Database.Open sSQL, oNewSecCon, adOpenKeyset, adLockPessimistic, adCmdText
            
            'Check to see if User/database combonation already exists in new UserDatabase table
            If rs30Database.RecordCount = 0 Then
            
                sSQL = "INSERT INTO UserDatabase (UserName, DatabaseCode) " _
                    & " VALUES ('" & sUsername & "','" & sDatabaseCode & "')"
                oNewSecCon.Execute sSQL
            
            End If
            
            rs22Database.MoveNext
        Loop
    End If
    
    
'************************* Role ******************************
Dim sRoleCode As String

    'Get all roles except MACROUser as this will already exist in the new Security DB
    sSQL = "SELECT * FROM Role WHERE RoleCode <> 'MACROUser'"
    Set rs22Database = New ADODB.Recordset
    rs22Database.Open sSQL, oOldSecCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    If rs22Database.RecordCount <> 0 Then
    
        Do While Not rs22Database.EOF
            
            sRoleCode = rs22Database!RoleCode
            
            'REM 31/07/03 - check in Uppercase as Oracle is case sensitive
            Select Case Connection_Property(CONNECTION_PROVIDER, sNewSecurityDBCon)
            Case CONNECTION_MSDAORA, CONNECTION_ORAOLEDB_ORACLE
                sSQL = "SELECT * FROM Role WHERE upper(RoleCode) = upper('" & sRoleCode & "')"
            Case Else
                sSQL = "SELECT * FROM Role WHERE RoleCode = '" & sRoleCode & "'"
            End Select

            Set rs30Database = New ADODB.Recordset
            rs30Database.Open sSQL, oNewSecCon, adOpenKeyset, adLockPessimistic, adCmdText
            
            If rs30Database.RecordCount = 0 Then
                
                sSQL = "INSERT INTO Role (ROLECODE,ROLEDESCRIPTION,ENABLED,SYSADMIN) " _
                    & " VALUES ('" & sRoleCode & "','" & rs22Database!roledescription & "'," _
                    & rs22Database!Enabled & "," & 0 & ")"
                oNewSecCon.Execute sSQL
                
            End If
            
            rs22Database.MoveNext
        Loop
    End If
    
'********************* Role Function ************************************
Dim sFunctionCode As String

    sSQL = "SELECT * FROM RoleFunction"
    Set rs22Database = New ADODB.Recordset
    rs22Database.Open sSQL, oOldSecCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    If rs22Database.RecordCount <> 0 Then
    
        Do While Not rs22Database.EOF
            
            sRoleCode = rs22Database!RoleCode
            sFunctionCode = rs22Database!FunctionCode
            
            sSQL = "SELECT * FROM RoleFunction "
            
            'REM 31/07/03 - check in Uppercase as Oracle is case sensitive
            Select Case Connection_Property(CONNECTION_PROVIDER, sNewSecurityDBCon)
            Case CONNECTION_MSDAORA, CONNECTION_ORAOLEDB_ORACLE
                sSQL = sSQL & "WHERE upper(RoleCode) = upper('" & sRoleCode & "')" _
                            & " AND FunctionCode = '" & sFunctionCode & "'"
            Case Else
                sSQL = sSQL & "WHERE RoleCode = '" & sRoleCode & "'" _
                            & " AND FunctionCode = '" & sFunctionCode & "'"
            End Select

            Set rs30Database = New ADODB.Recordset
            rs30Database.Open sSQL, oNewSecCon, adOpenKeyset, adLockPessimistic, adCmdText
            
            If rs30Database.RecordCount = 0 Then
                
                sSQL = "INSERT INTO RoleFunction (ROLECODE,FUNCTIONCODE) " _
                    & " VALUES ('" & sRoleCode & "','" & sFunctionCode & "')"
                oNewSecCon.Execute sSQL
            
            End If
        
            rs22Database.MoveNext
        Loop
    End If

'************** MACRO22UserRole Table *************************
Dim sUserNameDT As String
Dim sRoleCodeDT As String
Dim sDatabaseCodeDT As String
Dim sAllTrialsDT As String
Dim sAllSitesDT As String

    Select Case nDatabaseType
    Case MACRODatabaseType.sqlserver
        sUserNameDT = "VARCHAR(20)"
        sRoleCodeDT = "VARCHAR(15)"
        sDatabaseCodeDT = "VARCHAR(15)"
        sAllSitesDT = "SMALLINT"
        sAllTrialsDT = "SMALLINT"
    Case MACRODatabaseType.Oracle80
        sUserNameDT = "VARCHAR2(20)"
        sRoleCodeDT = "VARCHAR2(15)"
        sDatabaseCodeDT = "VARCHAR2(15)"
        sAllSitesDT = "NUMBER(6)"
        sAllTrialsDT = "NUMBER(6)"
    End Select

    'create a MACRO 2.2 UserRole Table
    sSQL = "CREATE Table MACRO22USERROLE(UserName " & sUserNameDT & "," _
        & " RoleCode " & sRoleCodeDT & "," _
        & " DatabaseCode " & sDatabaseCodeDT & "," _
        & " AllTrials " & sAllTrialsDT & "," _
        & " AllSites " & sAllSitesDT & "," _
        & " CONSTRAINT PKMACRO22UserRole PRIMARY KEY (UserName,RoleCode,DatabaseCode))"
    oNewSecCon.Execute sSQL
    
    'Copy MACRO 2.2 UserRole data into MACRO22UserRole table
    sSQL = "SELECT * FROM UserRole"
    Set rs22Database = New ADODB.Recordset
    rs22Database.Open sSQL, oOldSecCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    If rs22Database.RecordCount <> 0 Then
        
        Do While Not rs22Database.EOF
        
            sUsername = rs22Database!UserName
            sRoleCode = rs22Database!RoleCode
            sDatabaseCode = rs22Database!DatabaseCode
            
            'check to see if the user role already exits
            sSQL = "SELECT * FROM MACRO22UserRole "
            'REM 31/07/03 - check in Uppercase as Oracle is case sensitive
            Select Case Connection_Property(CONNECTION_PROVIDER, sNewSecurityDBCon)
            Case CONNECTION_MSDAORA, CONNECTION_ORAOLEDB_ORACLE
                sSQL = sSQL & " WHERE upper(Username) = upper('" & sUsername & "')" _
                            & " AND upper(RoleCode) = upper('" & sRoleCode & "')" _
                            & " AND upper(DatabaseCode) = upper('" & sDatabaseCode & "')"
            Case Else
                sSQL = sSQL & " WHERE Username = '" & sUsername & "'" _
                            & " AND RoleCode = '" & sRoleCode & "'" _
                            & " AND DatabaseCode = '" & sDatabaseCode & "'"
            End Select
    
            Set rs30Database = New ADODB.Recordset
            rs30Database.Open sSQL, oNewSecCon, adOpenKeyset, adLockPessimistic, adCmdText
            
            'if not then insert it into MACRO22UserRole table
            If rs30Database.RecordCount = 0 Then
                sSQL = "INSERT INTO MACRO22UserRole (UserName,RoleCode,DatabaseCode,AllTrials,AllSites)" _
                    & " VALUES ('" & sUsername & "','" & sRoleCode & "','" _
                    & sDatabaseCode & "'," & rs22Database!AllTrials & "," & rs22Database!AllSites & ")"
                oNewSecCon.Execute sSQL
                
            End If
            
            rs22Database.MoveNext
        Loop
    End If

    rs22Database.Close
    Set rs22Database = Nothing
    rs30Database.Close
    Set rs30Database = Nothing
    
    oNewSecCon.Close
    Set oNewSecCon = Nothing
    oOldSecCon.Close
    Set oOldSecCon = Nothing

    HourglassOff

Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|modNewDatabaseandTables.ImportSecurityData"
End Sub


