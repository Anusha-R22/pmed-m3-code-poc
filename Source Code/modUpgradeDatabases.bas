Attribute VB_Name = "modUpgradeDatabases"
 ''------------------------------------------------------------------------------'
' File:         modUpgradeDatabases.bas
' Copyright:    InferMed Ltd. 2000. All Rights Reserved
' Author:       Mo Morris, January 2000
' Purpose:      Contains  routines for upgrading Macro's DATA databases
'               and SECURITY databases from version 2.0.15 to the current version
'------------------------------------------------------------------------------'`
'   Revisions:
'   TA 11/8/2001: Started Code for Upgrading from 2.1 to 2.2
'   Mo Morris   6/9/01  Full database upgrade from end of 2.1 to beginning of 2.2,
'               a revised UpgradeDataDatabase2_2 now calls UpgradeData2_1to2_2_1, which
'               performs all of the Database Audit changes.
'   Mo Morris   17/10/01,   New sub DropDefaultConstraint added
'   NCJ 19 Oct 01 - Upgrade 2.2.1 to 2.2.2
'   rem 26/10/01 - Upgrade 2.2.2 to 2.2.3
'   rem 20/11/01 - Upgrade 2.2.3 to 2.2.4
'   TA 6/12/2001: Pass through connection object rather than string when creating KEYWORDS table so that MACROADODBConnection is up to date
'### WHEN CALLING CreateDB IN AN UPGRADE ALWAYS PASS THOUGH THE CONNECTION OBJECT ###
'   REM 18/04/02 - Upgrade 2.2.5 to 2.2.10
'   REM 07/06/02 - Changed DAO to ADO connections for some of the MACRO 2.1 database upgrade
'   TA 13/06/02 CBB 2.2.13.43, F5006 (view reports) function reinserted into Function and RoleFunction tables
'   TA 15/08/2002:  Preaparation done for MACROAccess22ToMSDE22 transfer - call to it commented out until upgrade is written
'   RS 16/9/2002:   Added UpgradeMACRO3_0from9to30 (TimezoneOffset columns)
'   NCJ 16 Oct 02 - Added Security 15 to 16
'   NCJ 20 Dec 02 - Upgrade for 3.0 Build 28
'   NCJ 2 Jan 03 - Upgrade for 3.0 Build 29
'   NCJ 30 Jan 03 - Upgrade for 3.0 Build 34 (first Roche version)
'------------------------------------------------------------------------------'

Option Explicit

Public Const CURRENT_SUBVERSION = 50  'change this no each build

'------------------------------------------------------------------------------'
Private Function CreateMacroTablev2_1_4(oExecuter As Object, bAccess As Boolean)
'------------------------------------------------------------------------------'
'TA 12/10/2000: New MACROTable table - stores all tables
' NCJ 8/11/00 - Include Units, ClinicalTest and ClinicalTestGroup in LDD
'------------------------------------------------------------------------------'
Dim sSQLStart As String
Dim sSQL As String
    On Error GoTo ErrHandler
    
    'TA 12/10/2000: New MACROTable table - stores all tables
    If bAccess Then
        'Using DAO
        sSQL = "CREATE Table MACROTable (TableName TEXT(30),"
        sSQL = sSQL & " SegmentId TEXT(3),"
        sSQL = sSQL & " STYDEF SMALLINT, PATRSP SMALLINT, LABDEF SMALLINT,"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (TableName))"
        'REM 07/06/02 - changed to use ADO connection
'        With oExecuter
'            .Execute sSQL, dbFailOnError
'            .TableDefs.Refresh
'        End With
        
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table MACROTable (TableName VARCHAR(30),"
        sSQL = sSQL & " SegmentId VARCHAR(3),"
        sSQL = sSQL & " STYDEF INTEGER, PATRSP INTEGER, LABDEF INTEGER,"
        sSQL = sSQL & " CONSTRAINT PKMACROTable PRIMARY KEY "
        sSQL = sSQL & " (TableName))"
        
    End If
    oExecuter.Execute sSQL
    
    sSQLStart = "INSERT INTO MACROTable VALUES ("
    
'3rd column for SDD, 4th for PRD, 5th for LDD
' NCJ 25/10/00 - Added new values for Units and UnitConversionFactors
' NCJ 8/11/00 - Include Units, ClinicalTest and ClinicalTestGroup in LDD
With oExecuter
        .Execute sSQLStart & "'ClinicalTest', '620', 1, 0, 1)"
        .Execute sSQLStart & "'ClinicalTestGroup', '610', 1, 0, 1)"
        .Execute sSQLStart & "'ClinicalTrial', '001', 1, 0, 0)"
        .Execute sSQLStart & "'CRFElement', '140', 1, 0, 0)"
        .Execute sSQLStart & "'CRFPage', '130', 1, 0, 0)"
        .Execute sSQLStart & "'CRFPageInstance', '070', 0, 1, 0)"
        .Execute sSQLStart & "'CTC', '640', 1, 0, 0)"
        .Execute sSQLStart & "'CTCScheme', '630', 1, 0, 0)"
        .Execute sSQLStart & "'DataItem', '120', 1, 0, 0)"
        .Execute sSQLStart & "'DataItemResponse', '080', 0, 1, 0)"
        .Execute sSQLStart & "'DataItemResponseHistory', '090', 0, 1, 0)"
        .Execute sSQLStart & "'DataItemValidation', '125', 1, 0, 0)"
        .Execute sSQLStart & "'DataType', '', 0, 0, 0)"
        .Execute sSQLStart & "'ExternalDataMapping', '', 0, 0, 0)"
        .Execute sSQLStart & "'ExternalDataSource', '', 0, 0, 0)"
        .Execute sSQLStart & "'Laboratory', '001', 0, 0, 1)"
        .Execute sSQLStart & "'LogDetails', '', 0, 0, 0)"
        .Execute sSQLStart & "'MACROControl', '', 0, 0, 0)"
        .Execute sSQLStart & "'MACROTable', '', 0, 0, 0)"
        .Execute sSQLStart & "'Message', '', 0, 0, 0)"
        .Execute sSQLStart & "'MIMessage', '', 0, 0, 0)"
        .Execute sSQLStart & "'NewDBColumn', '', 0, 0, 0)"
        .Execute sSQLStart & "'NormalRange', '030', 0, 0, 1)"
'        .Execute sSQLStart & "'PRDExportImport', '', 0, 0, 0)"
        .Execute sSQLStart & "'Protocols', '300', 1, 0, 0)"
        .Execute sSQLStart & "'RandomisationStep', '170', 1, 0, 0)"
        .Execute sSQLStart & "'ReasonForChange', '060', 1, 0, 0)"
        .Execute sSQLStart & "'ReportType', '', 0, 0, 0)"
        .Execute sSQLStart & "'RequiredData', '', 0, 0, 0)"
'        .Execute sSQLStart & "'SDDExportImport', '', 0, 0, 0)"
        .Execute sSQLStart & "'Site', '', 0, 0, 0)"
        .Execute sSQLStart & "'SiteLaboratory', '', 0, 0, 0)"
        .Execute sSQLStart & "'SiteUser', '', 0, 0, 0)"
        .Execute sSQLStart & "'StandardDataFormat', '', 0, 0, 0)"
        .Execute sSQLStart & "'StratificationFactor', '180', 1, 0, 0)"
        .Execute sSQLStart & "'StudyDefinition', '030', 1, 0, 0)"
        .Execute sSQLStart & "'StudyDocument', '040', 1, 0, 0)"
        .Execute sSQLStart & "'StudyReport', '190', 1, 0, 0)"
        .Execute sSQLStart & "'StudyReportData', '200', 1, 0, 0)"
        .Execute sSQLStart & "'StudyVisit', '150', 1, 0, 0)"
        .Execute sSQLStart & "'StudyVisitCRFPage', '160', 1, 0, 0)"
        .Execute sSQLStart & "'TrialOffice', '', 0, 0, 0)"
        .Execute sSQLStart & "'TrialPhase', '510', 1, 0, 0)"
        .Execute sSQLStart & "'TrialSite', '', 0, 0, 0)"
        .Execute sSQLStart & "'TrialStatus', '', 0, 0, 0)"
        .Execute sSQLStart & "'TrialStatusHistory', '050', 1, 0, 0)"
        .Execute sSQLStart & "'TrialSubject', '040', 0, 1, 0)"
        .Execute sSQLStart & "'TrialType', '520', 1, 0, 0)"
        .Execute sSQLStart & "'UnitClasses', '', 0, 0, 0)"
        .Execute sSQLStart & "'UnitConversionFactors', '560', 1, 0, 0)"
        .Execute sSQLStart & "'Units', '550', 1, 0, 1)"
        .Execute sSQLStart & "'ValidationAction', '530', 1, 0, 0)"
        .Execute sSQLStart & "'ValidationType', '540', 1, 0, 0)"
        .Execute sSQLStart & "'ValueData', '090', 1, 0, 0)"
        .Execute sSQLStart & "'VisitInstance', '060', 0, 1, 0)"
    End With

Exit Function
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CreateMacroTablev2_1_4", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function


'------------------------------------------------------------------------------'
Private Function CreateNewDBColumn2_1_4(oExecuter As Object, bAccess As Boolean)
'------------------------------------------------------------------------------'
'TA 12/10/2000: New CreateNewDBColumn table - stores added columns and the version it happened
'------------------------------------------------------------------------------'
Dim sSQLStart As String
Dim sSQL As String
    On Error GoTo ErrHandler
    
    'TA 12/10/2000: New MACROTable table - stores all tables
    If bAccess Then
        'Using DAO
        sSQL = "CREATE Table NewDBColumn (VersionMajor SMALLINT, VersionMinor SMALLINT, VersionRevision SMALLINT,"
        sSQL = sSQL & " TableName TEXT(30), ColumnName TEXT(30), ColumnOrder SMALLINT, DefaultValue TEXT(255),"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (VersionMajor, VersionMinor, VersionRevision, TableName, ColumnName))"
        'REM 07/06/02 - changed to ADO connection
'        With oExecuter
'            .Execute sSQL, dbFailOnError
'            .TableDefs.Refresh
'        End With
        
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table NewDBColumn (VersionMajor INTEGER, VersionMinor INTEGER, VersionRevision INTEGER,"
        sSQL = sSQL & " TableName VARCHAR(30), ColumnName VARCHAR(30), ColumnOrder INTEGER, DefaultValue VARCHAR(255),"
        sSQL = sSQL & " CONSTRAINT PKNewDBColumn PRIMARY KEY"
        sSQL = sSQL & " (VersionMajor, VersionMinor, VersionRevision, TableName, ColumnName))"

    End If
    oExecuter.Execute sSQL
    
    sSQLStart = "INSERT INTO NewDBColumn VALUES ("
    
    With oExecuter
        .Execute sSQLStart & "2,1,4,'CRFElement','ClinicalTestDate',null,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'DataItem','ClinicalTestDate',null,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'StudyDefinition','CTCSchemeCode',1,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'StudyDefinition','DOBExpr',2,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'StudyDefinition','GenderExpr',3,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'Trialsubject','SubjectGender',null,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'CRFPageInstance','LaboratoryCode',null,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'DataItemResponse','LabResult',1,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'DataItemResponse','CTCGrade',2,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'DataItemResponse','ClinicalTestDate',3,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'DataItemResponseHistory','LabResult',1,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'DataItemResponseHistory','CTCGrade',2,'#NULL#')"
        .Execute sSQLStart & "2,1,4,'DataItemResponseHistory','ClinicalTestDate',3,'#NULL#')"
    End With

Exit Function
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CreateNewDBColumn2_1_4", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function


'---------------------------------------------------------------------
Public Sub UpgradeSecurityDatabase()
'---------------------------------------------------------------------
' Check for necessary upgrading of the Security.mdb
' We do version 2.0 changes first, then 2.1
'---------------------------------------------------------------------

    ' Upgrade 2.0 Security database
    Call UpgradeSecurityDatabase2_0
    ' Upgrade 2.1 Security database
    Call UpGradeSecurityDatabase2_1
    ' Upgrade 2.2 Security database
    Call UpGradeSecurityDatabase2_2
    ' Upgrade 3.0 Security database
    Call UpGradeSecurityDatabase3_0

End Sub

'---------------------------------------------------------------------
Public Sub UpgradeDataDatabase()
'---------------------------------------------------------------------
' Check for necessary upgrading of the MACRO.mdb
' We do version 2.0 changes first, then 2.1
'---------------------------------------------------------------------

    ' Upgrade 2.0 Data database
    Call UpgradeDataDatabase2_0
    ' Upgrade 2.1 Data database
    Call UpgradeDataDatabase2_1
    ' Upgrade 2.2 Data database 'TA 11/8/2001
    Call UpgradeDataDatabase2_2
    
    'TODO: 3.0 Datasbe Upgrade TA 16/8/2002
    Call UpgradeDataDatabase3_0
    
End Sub

'---------------------------------------------------------------------
Private Sub UpGradeSecurityDatabase3_0()
'---------------------------------------------------------------------
' Upgrade a 2.2 Security database, checking first for upgrade from 2.0 to 2.1
'---------------------------------------------------------------------
Dim lBuildSubVersion As Long
Dim rsTemp As ADODB.Recordset
Dim sSQL As String
Dim sMacroVersion As String
Dim sBuildSubVersion As String
Dim sMsg As String
Dim sUpgradePath As String
Dim sScriptPrefix As String

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM SecurityControl"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        MsgBox ("Your Macro Security database is not valid. Macro is being closed down.")
        ExitMACRO
        MACROEnd
        Exit Sub
    End If
    
    ' Pick up the database version
    sMacroVersion = rsTemp![MACROVersion]
    sBuildSubVersion = rsTemp![BuildSubVersion]
    rsTemp.Close
    Set rsTemp = Nothing
    
    ' Check for version 2.1
    If sMacroVersion = "2.2" Then
        sMsg = "You are about to upgrade your MACRO security database from 2.2 to 3.0. Do you wish to continue?"
        Select Case MsgBox(sMsg, vbQuestion + vbYesNo, gsDIALOG_TITLE)
            Case vbYes
                sMacroVersion = "3.0"
                sBuildSubVersion = "7"
                Call UpGradeSecurity2_2to3_0_7
            Case vbNo
                Call ExitMACRO
                Call MACROEnd
        End Select
    End If
   
    If sMacroVersion <> "3.0" Then
        MsgBox ("Your MACRO Security database is not valid. MACRO is being closed down.")
        ExitMACRO
        MACROEnd
    End If
    
    'TODO remove this when we release
    ' anyhting below 7 needs manual upgrade
    If Val(sBuildSubVersion) < 7 Then
        DialogInformation "A manual upgrade is required for a " & sMacroVersion & "." & sBuildSubVersion & " security database"
    End If
    
    
'    'ADD NEW 3.0 UPGRADES HERE
'    If sBuildSubVersion = "x" Then
'        sBuildSubVersion = "x+1"
'        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
'    End If

    If Val(sBuildSubVersion) < 9 Then
        sBuildSubVersion = "9"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    'REM 08/10/02 - Upgrade from 9 to 14
    If Val(sBuildSubVersion) = 9 Then
        sBuildSubVersion = "14"
        Call UpgradeSecurity3_0from9to14
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'REM 14/10/02 - Upgrade from 14 to 15
    If Val(sBuildSubVersion) = 14 Then
        sBuildSubVersion = "15"
        Call UpgradeSecurity3_0from14to15
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    'NCJ 16/10/02 - Upgrade from 15 to 16
    If Val(sBuildSubVersion) = 15 Then
        sBuildSubVersion = "16"
        Call UpgradeSecurity3_0from15to16
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'ASH 21/10/02 - Upgrade from 16 to 18
    If Val(sBuildSubVersion) = 16 Then
        sBuildSubVersion = "18"
        Call UpgradeSecurity3_0from16to18
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    'REM 05/11/02 - Upgrade 18 to 20
    If (Val(sBuildSubVersion) >= 18) And (Val(sBuildSubVersion) <= 19) Then
        sBuildSubVersion = "20"
        Call UpgradeSecurity3_0from18to20
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'TA 12/11/02 - Upgrade from 20 to 21
    If Val(sBuildSubVersion) = 20 Then
        sBuildSubVersion = "21"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'TA 15/11/02 - Upgrade from 21 to 22
    If Val(sBuildSubVersion) = 21 Then
        sBuildSubVersion = "22"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'TA 18/11/02 - Upgrade from 22 to 23
    If Val(sBuildSubVersion) = 22 Then
        sBuildSubVersion = "23"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'TA 20/11/02 - Upgrade from 23 to 24
'these is no upgrade - security databases create anew


    'new method from version 24
    sUpgradePath = App.Path & "\Database Scripts\Upgrade Database\"
    
    'set prefix of cript file according to dbtype
    Select Case SecurityDatabaseType
    Case MACRODatabaseType.Oracle80: sScriptPrefix = "Security_ORA"
    Case Else: sScriptPrefix = "Security_MSSQL"
    End Select
    
    
    For lBuildSubVersion = 24 To (CURRENT_SUBVERSION - 1) 'change this no each build - always 1 less than build number
        If (Val(sBuildSubVersion) = lBuildSubVersion) Then
            sBuildSubVersion = CStr(lBuildSubVersion + 1)
            ExecuteMultiLineSQL SecurityADODBConnection, _
                                StringFromFile(sUpgradePath & sScriptPrefix & "_30_" & CStr(lBuildSubVersion) & "To" & sBuildSubVersion & ".sql")
        End If
    Next
    

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurityDatabase3_0", "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub
'---------------------------------------------------------------------
Private Sub UpgradeSecurity3_0from9to14()
'---------------------------------------------------------------------
'REM 08/10/02
'Add new table, PasswordHistory
'Add new columns to MACROUser: FailedAttempts, UserCreated
'                   MacroPassword: EnforceMixedCase, EnforceDigit, AllowRepeatChars, AllowUserName, PasswordHistory, PasswordRetries
'Add default values
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsPswd As ADODB.Recordset
Dim sInteger As String
Dim sText50 As String
Dim sText100 As String
Dim sDblDate As String
Dim sUserPswd As String
Dim sUsername As String
Dim sHashedPswd As String

    On Error GoTo ErrLabel

    'Create new Password History table
    sSQL = "CREATE Table PasswordHistory(UserName TEXT(50)," _
        & " HistoryNumber INTEGER," _
        & " UserCreated DOUBLE," _
        & " UserPassword TEXT(100)," _
        & " CONSTRAINT PKPasswordHistory PRIMARY KEY (UserName, HistoryNumber))"
    SecurityADODBConnection.Execute sSQL
    
    'add new columns to MACROUser
    sSQL = "ALTER Table MACROUser ADD COLUMN FailedAttempts INTEGER"
    SecurityADODBConnection.Execute sSQL
     sSQL = "ALTER Table MACROUser ADD COLUMN PasswordCreated DOUBLE"
    SecurityADODBConnection.Execute sSQL
    'add new columns to MacroPassword table
    sSQL = "ALTER Table MacroPassword ADD COLUMN EnforceMixedCase INTEGER"
    SecurityADODBConnection.Execute sSQL
    sSQL = "ALTER Table MacroPassword ADD COLUMN EnforceDigit INTEGER"
    SecurityADODBConnection.Execute sSQL
    sSQL = "ALTER Table MacroPassword ADD COLUMN AllowRepeatChars INTEGER"
    SecurityADODBConnection.Execute sSQL
    sSQL = "ALTER Table MacroPassword ADD COLUMN AllowUserName INTEGER"
    SecurityADODBConnection.Execute sSQL
    sSQL = "ALTER Table MacroPassword ADD COLUMN PasswordHistory INTEGER"
    SecurityADODBConnection.Execute sSQL
    sSQL = "ALTER Table MacroPassword ADD COLUMN PasswordRetries INTEGER"
    SecurityADODBConnection.Execute sSQL

    'add default values to new columns
    sSQL = "UPDATE MACROUser SET FailedAttempts = 0 WHERE FailedAttempts IS NULL"
    SecurityADODBConnection.Execute sSQL
    sSQL = "UPDATE MACROUser SET PasswordCreated = 0 WHERE PasswordCreated IS NULL"
    SecurityADODBConnection.Execute sSQL
    
    sSQL = "UPDATE MacroPassword SET EnforceMixedCase = 0 WHERE EnforceMixedCase IS NULL"
    SecurityADODBConnection.Execute sSQL
    sSQL = "UPDATE MacroPassword SET EnforceDigit = 0 WHERE EnforceDigit IS NULL"
    SecurityADODBConnection.Execute sSQL
    sSQL = "UPDATE MacroPassword SET AllowRepeatChars = 1 WHERE AllowRepeatChars IS NULL"
    SecurityADODBConnection.Execute sSQL
    sSQL = "UPDATE MacroPassword SET AllowUserName = 1 WHERE AllowUserName IS NULL"
    SecurityADODBConnection.Execute sSQL
    sSQL = "UPDATE MacroPassword SET PasswordHistory = 1 WHERE PasswordHistory IS NULL"
    SecurityADODBConnection.Execute sSQL
    sSQL = "UPDATE MacroPassword SET PasswordRetries = 5 WHERE PasswordRetries IS NULL"
    SecurityADODBConnection.Execute sSQL
    
    'Change Password field size in MACROUser table from text50 to text100 to hold new hashed password
    sSQL = "ALTER Table MACROUser ALTER Column UserPassword TEXT(100)"
    SecurityADODBConnection.Execute sSQL
    
    'Hash existing passswords
    sSQL = "SELECT UserName, UserPassword" _
        & " FROM MACROUser"
    Set rsPswd = New ADODB.Recordset
    rsPswd.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    Do While Not rsPswd.EOF
        sUserPswd = rsPswd!UserPassword
        sUsername = rsPswd!UserName
        
        sHashedPswd = HashHexEncodeString(sUserPswd)
        
        sSQL = "UPDATE MACROUser SET UserPassword = '" & sHashedPswd & "'" _
            & " WHERE UserName = '" & sUsername & "'"
        SecurityADODBConnection.Execute sSQL
    
        rsPswd.MoveNext
    Loop

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modUpgradeDatabases.UpGradeSecurity3_0from9to14"
End Sub

'---------------------------------------------------------------------
Private Sub UpgradeSecurity3_0from14to15()
'---------------------------------------------------------------------
'REM 14/10/02
'
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrLabel

    'Create new LoginLog table
    sSQL = "CREATE Table LoginLog(LogDateTime DOUBLE," _
        & " LogNumber INTEGER," _
        & " TaskId TEXT(50)," _
        & " LogMessage TEXT(255)," _
        & " UserName TEXT(20)," _
        & " CONSTRAINT PKLoginLog PRIMARY KEY (LogDateTime,LogNumber,TaskId))"
    SecurityADODBConnection.Execute sSQL

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modUpgradeDatabases.UpGradeSecurity3_0from14to15"
End Sub

'---------------------------------------------------------------------
Private Sub UpgradeSecurity3_0from15to16()
'---------------------------------------------------------------------
' NCJ 16/10/02
'
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrLabel

    ' New View SDV permission
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5023','View SDV mark')"
    SecurityADODBConnection.Execute sSQL

    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5023')"
    SecurityADODBConnection.Execute sSQL

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modUpgradeDatabases.UpGradeSecurity3_0from15to16"
End Sub

'---------------------------------------------------------------------
Private Sub UpGradeSecurity2_2to3_0_7()
'---------------------------------------------------------------------
' This upgrades a 2.2.x (latest) Security database to 3.0.7
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
  
    '*** Insert new Function Code ***
    sSQL = "INSERT INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F3025','Maintain question groups')"
    SecurityADODBConnection.Execute sSQL
    
    sSQL = "INSERT INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F3025')"
    SecurityADODBConnection.Execute sSQL
    
    'Upgrade MACROVersion from [2.2] to [3.0]
    sSQL = "UPDATE SecurityControl Set MACROVersion = '3.0'"
    SecurityADODBConnection.Execute sSQL

    'Upgrade BuildSubVersion to 1
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '7'"
    SecurityADODBConnection.Execute sSQL
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurity2_2to3_0_7", "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeSecurityDatabase2_2()
'---------------------------------------------------------------------
' Upgrade a 2.1 Security database, checking first for upgrade from 2.0 to 2.1
'---------------------------------------------------------------------

Dim rsTemp As ADODB.Recordset
Dim sSQL As String
Dim sMacroVersion As String
Dim sBuildSubVersion As String
Dim sMsg As String

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM SecurityControl"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        MsgBox ("Your Macro Security database is not valid. Macro is being closed down.")
        ExitMACRO
        MACROEnd
        Exit Sub
    End If
    
    ' Pick up the database version
    sMacroVersion = rsTemp![MACROVersion]
    sBuildSubVersion = rsTemp![BuildSubVersion]
    rsTemp.Close
    Set rsTemp = Nothing
    
    'don't do anything if futre version
    If Val(sMacroVersion) > Val(2.2) Then
        Exit Sub
    End If
    
    ' Check for version 2.1
    If sMacroVersion = "2.1" Then
        sMsg = "You are about to upgrade your MACRO security database from 2.1 to 2.2. Do you wish to continue?"
        Select Case MsgBox(sMsg, vbQuestion + vbYesNo, gsDIALOG_TITLE)
            Case vbYes
                sMacroVersion = "2.2"
                sBuildSubVersion = "1"
                Call UpGradeSecurity2_1to2_2_1
            Case vbNo
                Call ExitMACRO
                Call MACROEnd
        End Select
    End If
   
    If sMacroVersion <> "2.2" Then
        MsgBox ("Your MACRO Security database is not valid. MACRO is being closed down.")
        ExitMACRO
        MACROEnd
    End If
    
    'ADD NEW 2.2 UPGRADES HERE
    If sBuildSubVersion = "1" Then
        sBuildSubVersion = "2"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    'REM 26/10/01 - 2 to 3
    If sBuildSubVersion = "2" Then
        sBuildSubVersion = "3"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    'REM 20/11/01 - 3 to 4
    If sBuildSubVersion = "3" Then
        sBuildSubVersion = "4"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'REM 10/12/01 - 4 to 5
    If sBuildSubVersion = "4" Then
        sBuildSubVersion = "5"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'MLM 20/12/01 - 5 to 6
    If sBuildSubVersion = "5" Then
        sBuildSubVersion = "6"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    'REM 11/01/02 - 6 to 7
    If sBuildSubVersion = "6" Then
        sBuildSubVersion = "7"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'REM 18/01/02 - 7 to 8
    If sBuildSubVersion = "7" Then
        sBuildSubVersion = "8"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
        Call UpgradeSecurity2_2From7to8
    End If

    'REM 31/01/02 - 8 to 9
    If sBuildSubVersion = "8" Then
        sBuildSubVersion = "9"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    'REM 15/04/02 - 9 to 10
    If sBuildSubVersion = "9" Then
        sBuildSubVersion = "10"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
        Call UpgradeSecurity2_2From9to10
    End If

    'MLM 28/04/02 - 10 to 11
    If sBuildSubVersion = "10" Then
        sBuildSubVersion = "11"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    'MLM 28/04/02 - 11 to 12
    If sBuildSubVersion = "11" Then
        sBuildSubVersion = "12"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'REM 10/05/02 - 12 to 13
    If sBuildSubVersion = "12" Then
        sBuildSubVersion = "13"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
        Call UpgradeSecurity2_2From12to13
    End If
    
    'ZA 14/05/02 - 13 to 14
    If sBuildSubVersion = "13" Then
        sBuildSubVersion = "14"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
        Call UpgradeSecurity2_2From13to14
    End If
    
    'REM 10/06/02 - 14 to 15
    If sBuildSubVersion = "14" Then
        sBuildSubVersion = "15"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'Mo 11/06/02 - 15 to 16
    'MLM 25/06/02: Modified from 16 to 18 since 16 & 17 were used for patches
    If sBuildSubVersion = "15" Then
        sBuildSubVersion = "18"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
        Call UpgradeSecurity2_2From15to16
    End If
    
    'MLM 28/04/02 - 18 to 19
    If sBuildSubVersion = "18" Then
        sBuildSubVersion = "19"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurityDatabase2_2", "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpgradeSecurity2_2From7to8()
'---------------------------------------------------------------------
' REM 18/01/02
' Change the online support password to an encrypted one
'---------------------------------------------------------------------
Dim rsOnlineSupport As ADODB.Recordset

    On Error GoTo ErrHandler

    Set rsOnlineSupport = New ADODB.Recordset
    rsOnlineSupport.Open "Select SupportUserPassword from OnlineSupport where SupportUserName = 'INFERMED'" _
                , SecurityADODBConnection, adOpenForwardOnly, adLockPessimistic, adCmdText
    rsOnlineSupport.Fields(0).Value = Crypt("guido")
    rsOnlineSupport.Update
    rsOnlineSupport.Close
    Set rsOnlineSupport = Nothing
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpgradeSecurity2_2From7to8", "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpgradeSecurity2_2From9to10()
'---------------------------------------------------------------------
' REM 15/04/02
' Add new column to Security database Databases table called SecureHTMLLocation
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    'add column to Databases table
    sSQL = "ALTER Table Databases ADD COLUMN SecureHTMLLocation Text(255)"
    SecurityADODBConnection.Execute sSQL
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpgradeSecurity2_2From9to10", "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'---------------------------------------------------------------------
Private Sub UpgradeSecurity2_2From12to13()
'---------------------------------------------------------------------
'REM 10/05/02
'Change the permission called "Update Arezzo from GGB import" to
'"Update Arezzo from Clinical Gateway import"
'---------------------------------------------------------------------
Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    sSQL = "UPDATE Function SET Function = 'Update Arezzo from Clinical Gateway import'" _
    & " WHERE FunctionCode = 'F3023'"
    SecurityADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpgradeSecurity2_2From12to13", "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'---------------------------------------------------------------------
Private Sub UpgradeSecurity2_2From13to14()
'---------------------------------------------------------------------
'ZA 14/05/02
'delete F3015 (create report) and F5006 (view reports) function
'from Function and RoleFunction table
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "Delete from RoleFunction where FunctionCode ='F3015'"
    SecurityADODBConnection.Execute sSQL
    
    sSQL = "Delete from RoleFunction where FunctionCode ='F5006'"
    SecurityADODBConnection.Execute sSQL
    
    sSQL = "Delete from Function where FunctionCode ='F3015'"
    SecurityADODBConnection.Execute sSQL
    
    sSQL = "Delete from Function where FunctionCode ='F5006'"
    SecurityADODBConnection.Execute sSQL
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpgradeSecurity2_2From13to14", "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'---------------------------------------------------------------------
Private Sub UpGradeSecurity2_1to2_2_1()
'---------------------------------------------------------------------
' This upgrades a 2.1.x (latest) Security database to 2.2.1
'---------------------------------------------------------------------
Dim sSQL As String
Dim oCat As Catalog

    On Error GoTo ErrHandler
  
    ' *** Drop Columns ***
    sSQL = "ALTER Table Databases DROP Column DocumentLocation"
    SecurityADODBConnection.Execute sSQL
    sSQL = "ALTER Table Databases DROP Column InFolderLocation"
    SecurityADODBConnection.Execute sSQL
    sSQL = "ALTER Table Databases DROP Column OutFolderLocation"
    SecurityADODBConnection.Execute sSQL
    sSQL = "ALTER Table Databases DROP Column TempLocation"
    SecurityADODBConnection.Execute sSQL
    'drop index on field SubSystem prior to dropping the column
    sSQL = "ALTER Table Function DROP CONSTRAINT idx_SubSystem"
    SecurityADODBConnection.Execute sSQL
    sSQL = "ALTER Table Function DROP Column SubSystem"
    SecurityADODBConnection.Execute sSQL
    'drop index on field ClinicalTrialId prior to dropping the column
    sSQL = "ALTER Table UserRole DROP CONSTRAINT idx_ClinicalTrialId "
    SecurityADODBConnection.Execute sSQL
    sSQL = "ALTER Table UserRole DROP Column ClinicalTrialId"
    SecurityADODBConnection.Execute sSQL
    sSQL = "ALTER Table UserRole DROP Column TrialSite"
    SecurityADODBConnection.Execute sSQL
    
    ' *** Foreign Keys ***
    'The following calls to FieldNameTypeUpgrade upgrade tables by first dropping them and then re-creating them.
    'Unfortunately the security database contains tables with foreign keys that prevent tables being dropped.
    'To get over this problem the foreign keys will be removed prior to the calls to FieldNameTypeUpgrade.
    'Note that FieldNameTypeUpgrade calls CreateDB to re-create the tables, with the foreign keys back in place.
    sSQL = "ALTER Table UserDatabase DROP CONSTRAINT FKUserdatabaseDD"
    SecurityADODBConnection.Execute sSQL
    sSQL = "ALTER Table UserRole DROP CONSTRAINT FKUserRoleDD"
    SecurityADODBConnection.Execute sSQL
    sSQL = "ALTER Table UserDatabase DROP CONSTRAINT FKUserdatabaseUC"
    SecurityADODBConnection.Execute sSQL
    sSQL = "ALTER Table UserRole DROP CONSTRAINT FKUserRoleUC"
    SecurityADODBConnection.Execute sSQL
    
    ' *** Columns with Name and/or Type Changes ***
    'Table Databases, field DatabaseDescription (Text(15)) changed to DatabaseCode (Text(15))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Security, goUser.Database.DatabaseType, SecurityADODBConnection, "Databases", "2.2.1")
    
    'Table MACROUser, field Password (Text(50)) changed to UserPassword (Text(50))
    'Table MACROUser, field UserCode (Text(50)) changed to UserName (Text(20))
    'Table MACROUser, field UserName (Text(50)) changed to UserNameFull (Text(100))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Security, goUser.Database.DatabaseType, SecurityADODBConnection, "MACROUser", "2.2.1")
    
    'Table UserDatabase, field DatabaseDescription (Text(15)) changed to DatabaseCode (Text(15))
    'Table UserDatabase, field UserCode(Text(50)) changed to UserName (Text(20))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Security, goUser.Database.DatabaseType, SecurityADODBConnection, "UserDatabase", "2.2.1")
    
    'Table UserRole, field DatabaseDescription (Text(15)) changed to DatabaseCode (Text(15))
    'Table UserRole, field UserCode(Text(50)) changed to UserName (Text(20))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Security, goUser.Database.DatabaseType, SecurityADODBConnection, "UserRole", "2.2.1")
    
    'Insert New Functions
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5021','Remove own locks')"
    SecurityADODBConnection.Execute sSQL

    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5021')"
    SecurityADODBConnection.Execute sSQL
    
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5022','Remove all locks')"
    SecurityADODBConnection.Execute sSQL

    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5022')"
    SecurityADODBConnection.Execute sSQL
    
    'Upgrade MACROVersion from [2.1] to [2.2]
    sSQL = "UPDATE SecurityControl Set MACROVersion = '2.2'"
    SecurityADODBConnection.Execute sSQL

    'Upgrade BuildSubVersion to 1
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '1'"
    SecurityADODBConnection.Execute sSQL

'    'change password column to UserPassword in MACROUser table
'    Set oCat = New Catalog
'    Set oCat.ActiveConnection = SecurityADODBConnection
'    oCat.Tables("MACROUser").Columns("Password").Name = "UserPassword"
'    Set oCat = Nothing

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurity2_1to2_2_1", "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub


'---------------------------------------------------------------------
Private Sub UpGradeSecurityDatabase2_1()
'---------------------------------------------------------------------
' Upgrade a 2.1 Security database, checking first for upgrade from 2.0 to 2.1
'---------------------------------------------------------------------

Dim rsTemp As ADODB.Recordset
Dim sSQL As String
Dim sMacroVersion As String
Dim sBuildSubVersion As String
Dim sMsg As String

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM SecurityControl"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        MsgBox ("Your Macro Security database is not valid. Macro is being closed down.")
        ExitMACRO
        MACROEnd
        Exit Sub
    End If
    
    ' Pick up the database version
    sMacroVersion = rsTemp![MACROVersion]
    sBuildSubVersion = rsTemp![BuildSubVersion]
    rsTemp.Close
    Set rsTemp = Nothing
    
    'don't do anything if futre version
    If Val(sMacroVersion) >= Val(2.2) Then
        Exit Sub
    End If
    ' Check for version 2.0
    If sMacroVersion = "2.0" Then
        'This is a temporary measure and must be taken out on everybody changing to 2.1
        sMsg = "You are about to upgrade your MACRO security database from 2.0 to 2.1. Do you wish to continue?"
        Select Case MsgBox(sMsg, vbQuestion + vbYesNo, gsDIALOG_TITLE)
            Case vbYes
                ' Do initial upgrade from 2.0.x to 2.1.1
                UpGradeSecurity2_0to2_1_1
                sMacroVersion = "2.1"
                sBuildSubVersion = "1"
            Case vbNo
                Call ExitMACRO
                Call MACROEnd
        End Select
    End If
   
    

    
    If sMacroVersion <> "2.1" Then
        MsgBox ("Your Macro Security database is not valid. Macro is being closed down.")
        ExitMACRO
        MACROEnd
    End If
    
    'ADD NEW 2.1 UPGRADES HERE
    'Check for "01" as well as "1" only necessary for Sub Version 1
    If sBuildSubVersion = "1" Or sBuildSubVersion = "01" Then
        UpGradeSecurity2_1from1to2
        sBuildSubVersion = "2"
    End If
   
    If sBuildSubVersion = "2" Then
        sBuildSubVersion = "3"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    If sBuildSubVersion = "3" Then
        sBuildSubVersion = "4"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    If sBuildSubVersion = "4" Or sBuildSubVersion = "5" Then
        ' We missed out 5
        sBuildSubVersion = "6"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
   
   ' NCJ 1/11/00 - Upgrade 6 to 11
   ' NCJ 28/11/00 - Include 7,8,9,10,11,12 as well!
    Select Case sBuildSubVersion
    Case "6", "7", "8", "9", "10", "11", "12"
        sBuildSubVersion = "13"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    Case Else
        ' Nothing to do here
    End Select
    
    'Mo Morris 29/11/00 - Upgrade 13 to 14
    If sBuildSubVersion = "13" Then
        UpGradeSecurity2_1from13to14
        sBuildSubVersion = "14"
    End If
    
    'Mo Morris 19/12/00 - Upgrade 14,15,16 to 17
    Select Case sBuildSubVersion
    Case "14", "15", "16"
        sBuildSubVersion = "17"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    Case Else
        ' Nothing to do here
    End Select
    
    'TA 30/1/2001: Upgrade from 17 to 30 ' leaving room for Ronald builds
    If Val(sBuildSubVersion) >= 17 And Val(sBuildSubVersion) < 30 Then
        sBuildSubVersion = "30"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    ' NCJ 8 Feb 01 Upgrade to 32
    ' NCJ 19 Feb 01 - Upgrade 30-33 to 34
    ' NCJ 8 Mar 01 - Upgrade 34 and 35 to 36
   ' NCJ 23/3/01 - Upgrade to 37
   ' NCJ 2/4/01 - Upgrade to 38
   ' NCJ 10/4/01 - Upgrade to 39
    If Val(sBuildSubVersion) >= 30 And Val(sBuildSubVersion) <= 38 Then
        sBuildSubVersion = "39"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'TA 18/04/2001 - Upgrade 39 to 40
    If sBuildSubVersion = "39" Then
        UpGradeSecurity2_1from39to40
        sBuildSubVersion = "40"
    End If


Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurityDatabase2_1", "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

   

End Sub

'---------------------------------------------------------------------
Private Sub UpgradeSecurityDatabase2_0()
'---------------------------------------------------------------------
' Upgrade a 2.0 Security.mdb
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim sSQL As String
Dim sMacroVersion As String
Dim sBuildSubVersion As String

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM SecurityControl"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        MsgBox ("Your Macro Security database is not valid. Macro is being closed down.")
        ExitMACRO
        MACROEnd
        Exit Sub
    End If
    
    sMacroVersion = rsTemp![MACROVersion]
    If sMacroVersion <> "2.0" Then
        Exit Sub
    End If
    
    sBuildSubVersion = rsTemp![BuildSubVersion]
    
    'Check for need to upgrade from BuildSubVersion [15] to [16]
    If sBuildSubVersion = "15" Then
        UpGradeSecurity15to16
        sBuildSubVersion = "16"
    End If
    
    'Check for need to upgrade from BuildSubVersion [16] to [17]
    If sBuildSubVersion = "16" Then
        sBuildSubVersion = "17"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'Check for need to upgrade from BuildSubVersion [17] to [18]
    If sBuildSubVersion = "17" Then
        sBuildSubVersion = "18"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'Check for need to upgrade from BuildSubVersion [18] to [19]
    If sBuildSubVersion = "18" Then
        sBuildSubVersion = "19"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'Check for need to upgrade from BuildSubVersion [19] to [20]
    If sBuildSubVersion = "19" Then
        UpGradeSecurity19to20   ' Special upgrade
        sBuildSubVersion = "20"
    End If

    'Check for need to upgrade from BuildSubVersion [20] to [21]
    If sBuildSubVersion = "20" Then
        sBuildSubVersion = "21"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    'Check for need to upgrade from BuildSubVersion [21] to [22]
    ' NCJ 8/3/00
    If sBuildSubVersion = "21" Then
        sBuildSubVersion = "22"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'Check for need to upgrade from BuildSubVersion [22] to [23]
    ' WIllC 14/3/00
    If sBuildSubVersion = "22" Then
        sBuildSubVersion = "23"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'Check for need to upgrade from BuildSubVersion [23] to [27]
    ' NCJ 3/4/00
    If sBuildSubVersion = "23" Then
        sBuildSubVersion = "27"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    'Check for need to upgrade from BuildSubVersion [27] to [29]
    ' WIllC 26/4/00
    If sBuildSubVersion = "27" Then
        UpGradeSecurity27to29
        sBuildSubVersion = "29"
    End If
    
    If sBuildSubVersion = "29" Then
        sBuildSubVersion = "31"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    If sBuildSubVersion = "31" Then
        sBuildSubVersion = "32"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    'WillC 16/5/00
    If sBuildSubVersion = "32" Then
        UpGradeSecurity32to33
        sBuildSubVersion = "33"
    End If

    'WillC 26/5/00
    If sBuildSubVersion = "33" Then
        UpGradeSecurity33to34
        sBuildSubVersion = "34"
    End If

    'WillC 30/5/00
    If sBuildSubVersion = "34" Then
        UpGradeSecurity34to35
        sBuildSubVersion = "35"
    End If
    
    'NCJ 2/6/00
    If sBuildSubVersion = "35" Then
        sBuildSubVersion = "36"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
        'WillC 9/6/00
    If sBuildSubVersion = "36" Then
        UpGradeSecurity36to37
        sBuildSubVersion = "37"
    End If
    
        'Nicky 15/6/00
    If sBuildSubVersion = "37" Then
        sBuildSubVersion = "39"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

        'Nicky 16/6/00
    If sBuildSubVersion = "39" Then
        sBuildSubVersion = "40"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpgradeSecurityDatabase", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub


'---------------------------------------------------------------------
Private Sub UpGradeData2_1to2_2_1()
'---------------------------------------------------------------------
' Upgrade 2.1.x (latest) to 2.2.1 (Includes all of the 2.1 to 2.2 database audit changes
'---------------------------------------------------------------------
Dim oCat As Catalog
Dim sSQL As String
Dim sDefaultConstraintName As String

    On Error GoTo ErrHandler
    
    ' *** Drop Tables ***
    sSQL = "DROP Table ExternalDataMapping"
    MacroADODBConnection.Execute sSQL
    sSQL = "DROP Table ExternalDataSource"
    MacroADODBConnection.Execute sSQL
    sSQL = "DROP Table RandomisationStep"
    MacroADODBConnection.Execute sSQL
    sSQL = "DROP Table ReportType"
    MacroADODBConnection.Execute sSQL
    sSQL = "DROP Table StratificationFactor"
    MacroADODBConnection.Execute sSQL
    
    ' *** Remove Dropped tables from table MACROTable ***
    sSQL = "DELETE  FROM MACROTable WHERE TableName = 'ExternalDataMapping'"
    MacroADODBConnection.Execute sSQL
    sSQL = "DELETE  FROM MACROTable WHERE TableName = 'ExternalDataSource'"
    MacroADODBConnection.Execute sSQL
    sSQL = "DELETE  FROM MACROTable WHERE TableName = 'RandomisationStep'"
    MacroADODBConnection.Execute sSQL
    sSQL = "DELETE  FROM MACROTable WHERE TableName = 'ReportType'"
    MacroADODBConnection.Execute sSQL
    sSQL = "DELETE  FROM MACROTable WHERE TableName = 'StratificationFactor'"
    MacroADODBConnection.Execute sSQL
    
    ' *** Add New Tables ***
    'Mo Morris 1/10/01, optional field to prevent CreateDB displaying messages added
    Call CreateDB(eMACRODBFunction.Data, goUser.Database.DatabaseType, MacroADODBConnection.ConnectionString, False, False, "MACROLock", "2.2.1")
    
    ' *** Add New Tables to table MACROTable ***
    sSQL = "INSERT INTO MACROTable VALUES ('MACROLock', '', 0, 0, 0)"
    MacroADODBConnection.Execute sSQL
    
    ' *** New Columns ***
    'Add eFormDatePrompt to table CRFPage
    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.Access
        sSQL = "ALTER Table CRFPage ADD COLUMN eFormDatePrompt BYTE"
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
        sSQL = "ALTER Table CRFPage ADD eFormDatePrompt TINYINT"
    Case MACRODatabaseType.Oracle80
        sSQL = "ALTER Table CRFPage ADD eFormDatePrompt NUMBER(3)"
    End Select
    MacroADODBConnection.Execute sSQL
    
    'Add ChangeType and ColumnNumber to table NewDBColumn
    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.Access
        sSQL = "ALTER Table NewDBColumn ADD COLUMN ChangeType TEXT(15)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table NewDBColumn ADD COLUMN ColumnNumber SMALLINT"
        MacroADODBConnection.Execute sSQL
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
        sSQL = "ALTER Table NewDBColumn ADD ChangeType VARCHAR(15)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table NewDBColumn ADD ColumnNumber SMALLINT"
        MacroADODBConnection.Execute sSQL
    Case MACRODatabaseType.Oracle80
        sSQL = "ALTER Table NewDBColumn ADD ChangeType VARCHAR2(15)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table NewDBColumn ADD ColumnNumber NUMBER(6)"
        MacroADODBConnection.Execute sSQL
    End Select
    
    'As a one off set all the existing entries in table NewDBColumn to have
    'their new field "ChangeType" set to "NEWCOLUMN"
    sSQL = "UPDATE NewDBColumn SET ChangeType = 'NEWCOLUMN'"
    MacroADODBConnection.Execute sSQL
    
    ' *** Add new columns to NewDBColumn ***
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'CRFPage','eFormDatePrompt',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'NewDBColumn','ChangeType',1,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'NewDBColumn','ColumnNumber',2,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    
    ' *** Drop Columns ***
    sSQL = "ALTER Table CRFElement DROP Column DefaultValue"
    MacroADODBConnection.Execute sSQL
    
    'Drop the Default value constraint before dropping the column in SQL Server databases
    If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Or goUser.Database.DatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("CRFPageInstance", "SDVStatus")
    End If
    sSQL = "ALTER Table CRFPageInstance DROP Column SDVStatus"
    MacroADODBConnection.Execute sSQL
    
    If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Or goUser.Database.DatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("CRFPageInstance", "ReviewStatus")
    End If
    sSQL = "ALTER Table CRFPageInstance DROP Column ReviewStatus"
    MacroADODBConnection.Execute sSQL
    
    sSQL = "ALTER Table DataItem DROP Column ExternalDataSource"
    MacroADODBConnection.Execute sSQL
    
    If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Or goUser.Database.DatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("DataItem", "RequiredTrialTypeId")
    End If
    sSQL = "ALTER Table DataItem DROP Column RequiredTrialTypeId"
    MacroADODBConnection.Execute sSQL
    
    If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Or goUser.Database.DatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("DataItem", "Required")
    End If
    sSQL = "ALTER Table DataItem DROP Column Required"
    MacroADODBConnection.Execute sSQL
    
    If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Or goUser.Database.DatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("DataItemResponse", "SDVStatus")
    End If
    sSQL = "ALTER Table DataItemResponse DROP Column SDVStatus"
    MacroADODBConnection.Execute sSQL
    
    If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Or goUser.Database.DatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("DataItemResponse", "ReviewStatus")
    End If
    sSQL = "ALTER Table DataItemResponse DROP Column ReviewStatus"
    MacroADODBConnection.Execute sSQL
    
    sSQL = "ALTER Table DataItemResponse DROP Column ReviewComment"
    MacroADODBConnection.Execute sSQL
    
    If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Or goUser.Database.DatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("DataItemResponseHistory", "SDVStatus")
    End If
    sSQL = "ALTER Table DataItemResponseHistory DROP Column SDVStatus"
    MacroADODBConnection.Execute sSQL
    
    If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Or goUser.Database.DatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("DataItemResponseHistory", "ReviewStatus")
    End If
    sSQL = "ALTER Table DataItemResponseHistory DROP Column ReviewStatus"
    MacroADODBConnection.Execute sSQL
    
    sSQL = "ALTER Table DataItemResponseHistory DROP Column ReviewComment"
    MacroADODBConnection.Execute sSQL
    
    If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Or goUser.Database.DatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("MACROControl", "DateReference")
    End If
    sSQL = "ALTER Table MACROControl DROP Column DateReference"
    MacroADODBConnection.Execute sSQL
    
    If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Or goUser.Database.DatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("StudyReport", "HideBlankRows")
    End If
    sSQL = "ALTER Table StudyReport DROP Column HideBlankRows"
    MacroADODBConnection.Execute sSQL
    
    If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Or goUser.Database.DatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("StudyReportData", "HideBlankRows")
    End If
    sSQL = "ALTER Table StudyReportData DROP Column HideBlankRows"
    MacroADODBConnection.Execute sSQL
    
    sSQL = "ALTER Table TrialOffice DROP Column TransferSchedule"
    MacroADODBConnection.Execute sSQL
    
    If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Or goUser.Database.DatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("TrialSubject", "SDVStatus")
    End If
    sSQL = "ALTER Table TrialSubject DROP Column SDVStatus"
    MacroADODBConnection.Execute sSQL
    
    If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Or goUser.Database.DatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("TrialSubject", "ReviewStatus")
    End If
    sSQL = "ALTER Table TrialSubject DROP Column ReviewStatus"
    MacroADODBConnection.Execute sSQL
    
    sSQL = "ALTER Table TrialSubject DROP Column Checksum"
    MacroADODBConnection.Execute sSQL
    
    sSQL = "ALTER Table TrialSubject DROP Column Initials"
    MacroADODBConnection.Execute sSQL
    
    sSQL = "ALTER Table TrialSubject DROP Column Surname"
    MacroADODBConnection.Execute sSQL
    
    sSQL = "ALTER Table TrialSubject DROP Column Forename"
    MacroADODBConnection.Execute sSQL
    
    sSQL = "ALTER Table TrialSubject DROP Column SubjectIdentificationCode"
    MacroADODBConnection.Execute sSQL
    
    sSQL = "ALTER Table UnitConversionFactors DROP Column ConversionExpression"
    MacroADODBConnection.Execute sSQL
    
    If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Or goUser.Database.DatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("ValueData", "DefaultCat")
    End If
    sSQL = "ALTER Table ValueData DROP Column DefaultCat"
    MacroADODBConnection.Execute sSQL
    
    If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Or goUser.Database.DatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("VisitInstance", "SDVStatus")
    End If
    sSQL = "ALTER Table VisitInstance DROP Column SDVStatus"
    MacroADODBConnection.Execute sSQL
    
    If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Or goUser.Database.DatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("VisitInstance", "ReviewStatus")
    End If
    sSQL = "ALTER Table VisitInstance DROP Column ReviewStatus"
    MacroADODBConnection.Execute sSQL
    
    ' *** Add dropped columns to NewDBColumn ***
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'CRFElement','DefaultValue',null,'#NULL#','DROPCOLUMN',28)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'CRFPageInstance','SDVStatus',1,'#NULL#','DROPCOLUMN',15)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'CRFPageInstance','ReviewStatus',2,'#NULL#','DROPCOLUMN',14)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'DataItem','ExternalDataSource',1,'#NULL#','DROPCOLUMN',19)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'DataItem','RequiredTrialTypeId',2,'#NULL#','DROPCOLUMN',16)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'DataItem','Required',3,'#NULL#','DROPCOLUMN',15)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'DataItemResponse','SDVStatus',1,'#NULL#','DROPCOLUMN',25)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'DataItemResponse','ReviewStatus',2,'#NULL#','DROPCOLUMN',24)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'DataItemResponse','ReviewComment',3,'#NULL#','DROPCOLUMN',18)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'DataItemResponseHistory','SDVStatus',1,'#NULL#','DROPCOLUMN',25)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'DataItemResponseHistory','ReviewStatus',2,'#NULL#','DROPCOLUMN',24)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'DataItemResponseHistory','ReviewComment',3,'#NULL#','DROPCOLUMN',18)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'MACROControl','DateReference',null,'#NULL#','DROPCOLUMN',2)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'StudyReport','HideBlankRows',null,'#NULL#','DROPCOLUMN',4)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'StudyReportData','HideBlankRows',null,'#NULL#','DROPCOLUMN',4)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'TrialOffice','TransferSchedule',null,'#NULL#','DROPCOLUMN',7)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'TrialSubject','SDVStatus',1,'#NULL#','DROPCOLUMN',18)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'TrialSubject','ReviewStatus',2,'#NULL#','DROPCOLUMN',17)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'TrialSubject','Checksum',3,'#NULL#','DROPCOLUMN',15)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'TrialSubject','Initials',4,'#NULL#','DROPCOLUMN',7)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'TrialSubject','Surname',5,'#NULL#','DROPCOLUMN',6)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'TrialSubject','Forename',6,'#NULL#','DROPCOLUMN',5)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'TrialSubject','SubjectIdentificationCode',7,'#NULL#','DROPCOLUMN',4)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'UnitConversionFactors','ConversionExpression',null,'#NULL#','DROPCOLUMN',5)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'ValueData','DefaultCat',null,'#NULL#','DROPCOLUMN',9)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'VisitInstance','SDVStatus',1,'#NULL#','DROPCOLUMN',12)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,1,'VisitInstance','ReviewStatus',2,'#NULL#','DROPCOLUMN',11)"
    MacroADODBConnection.Execute sSQL

    ' *** Columns with Name and/or Type Changes ***
    'Table CRFElement, field Local (Long) changed to LocalFlag (integer)
    'Table CRFElement, field Hidden (Long) changed to Hidden (integer)
    'Table CRFElement, field Mandatory (Long) changed to Mandatory (integer)
    'Table CRFElement, field Optional (Long) changed to Optional (integer)
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, goUser.Database.DatabaseType, MacroADODBConnection, "CRFElement", "2.2.1")
    
    'Table MIMessage, field MIMessageUserCode (Text(50)) changed to MIMessageUserName (Text(20))
    'Table MIMessage, field MIMessageUserName (Text(255)) changed to MIMessageUserNameFull (Text(100))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, goUser.Database.DatabaseType, MacroADODBConnection, "MIMessage", "2.2.1")
    
    'Table DataItemResponse, field UserId (Text(15)) changed to UserName (Text(20))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, goUser.Database.DatabaseType, MacroADODBConnection, "DataItemResponse", "2.2.1")
    
    'Table DataItemResponseHistory, field UserId (Text(15)) changed to UserName (Text(20))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, goUser.Database.DatabaseType, MacroADODBConnection, "DataItemResponseHistory", "2.2.1")
    
    'Table LogDetails, field UserId (Text(15)) changed to UserName (Text(20))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, goUser.Database.DatabaseType, MacroADODBConnection, "LogDetails", "2.2.1")
    
    'Table Message, field UserId (Text(15)) changed to UserName (Text(20))
    'Table Message, field MessageId is no longer an autonumber
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, goUser.Database.DatabaseType, MacroADODBConnection, "Message", "2.2.1")
    
    'Table SiteUser, field UserCode (Text(15)) changed to UserName (Text(20))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, goUser.Database.DatabaseType, MacroADODBConnection, "SiteUser", "2.2.1")
    
    'Table StudyDefinition, field UserId (Text(15)) changed to UserName (Text(20))
    'Table StudyDefinition, field RRUserName (Text(50) changed to RRUserName (Text(20)
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, goUser.Database.DatabaseType, MacroADODBConnection, "StudyDefinition", "2.2.1")
    
    'Table TrialStatusHistory, field UserId (Text(15)) changed to UserName (Text(20))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, goUser.Database.DatabaseType, MacroADODBConnection, "TrialStatusHistory", "2.2.1")
    
    'Table TrialOffice, field Site (Text(50)) changed to Site (Text(8))
    'Table TrialOffice, field Password (Text(50)) changed to UserPassword (Text(50))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Data, goUser.Database.DatabaseType, MacroADODBConnection, "TrialOffice", "2.2.1")
    
    'Upgrade MacroVersion to [2.2]
    sSQL = "UPDATE MACROControl Set MacroVersion = '2.2'"
    MacroADODBConnection.Execute sSQL
    
    'Upgrade BuildSubVersion to [1]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '1'"
    MacroADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeData2_1to2_2_1", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub



'---------------------------------------------------------------------
Private Sub UpgradeDataDatabase2_2()
'---------------------------------------------------------------------
' Upgrade a 2.2 MACRO.mdb, checking first for 2.1 -> 2.2
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim sSQL As String
Dim sMacroVersion As String
Dim sBuildSubVersion As String
Dim sMsg As String

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM MACROControl"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        MsgBox ("Your Macro database is not valid. Macro is being closed down.")
        ExitMACRO
        MACROEnd
    End If
    
    sBuildSubVersion = rsTemp![BuildSubVersion]
    sMacroVersion = rsTemp![MACROVersion]
    rsTemp.Close
    Set rsTemp = Nothing

    'don't do anything if future version
    If Val(sMacroVersion) > Val(2.2) Then
        Exit Sub
    End If

    ' Check for version 2.1
    If sMacroVersion = "2.1" Then
        'REM 31/07/03 - no longer want it to pop up this message
        'sMsg = "You are about to upgrade your MACRO database from 2.1 to 2.2. Do you wish to continue?"
        'Select Case MsgBox(sMsg, vbQuestion + vbYesNo, gsDIALOG_TITLE)
            'Case vbYes
                Call UpGradeData2_1to2_2_1
                sMacroVersion = "2.2"
                sBuildSubVersion = "1"
            'Case vbNo
                'Call ExitMACRO
                'Call MACROEnd
        'End Select
    End If
    
    If sMacroVersion <> "2.2" Then
        MsgBox ("Your MACRO database is not valid. MACRO is being closed down.")
        ExitMACRO
        MACROEnd
    End If
    
    'ADD NEW 2.2 UPGRADES HERE
    ' NCJ 19 Oct 01 - 1 to 2
    If sBuildSubVersion = "1" Then
        sBuildSubVersion = "2"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'REM 26/10/01 - 2 to 3
    If sBuildSubVersion = "2" Then
        sBuildSubVersion = "3"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'REM 20/11/01 - 3 to 4
    If sBuildSubVersion = "3" Then
        sBuildSubVersion = "4"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'RJCW 26/11/01 - 4 to 5
    If sBuildSubVersion = "4" Then
        sBuildSubVersion = "5"
        Call UpGradeData2_2from1to5
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'MLM 20/12/01 - 5 to 6
    If sBuildSubVersion = "5" Then
        sBuildSubVersion = "6"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'REM 11/01/02 - 6 to 7
    If sBuildSubVersion = "6" Then
        sBuildSubVersion = "7"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'REM 18/01/02 - 7 to 8
    If sBuildSubVersion = "7" Then
        sBuildSubVersion = "8"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'REM 31/01/02 - 8 to 9
    If sBuildSubVersion = "8" Then
        sBuildSubVersion = "9"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'REM 15/04/02 - 9 to 10
    If sBuildSubVersion = "9" Then
        sBuildSubVersion = "10"
        Call UpGradeData2_2from5to10
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'MLM 28/04/02 - 10 to 11
    If sBuildSubVersion = "10" Then
        sBuildSubVersion = "11"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'MLM 30/04/02 - 11 to 12
    If sBuildSubVersion = "11" Then
        sBuildSubVersion = "12"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'MLM 24/05/02 - 12 to 13
    If sBuildSubVersion = "12" Then
        sBuildSubVersion = "13"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'MLM 24/05/02 - 13 to 14 (whoops, forgot 13 upgrade)
    If sBuildSubVersion = "13" Then
        sBuildSubVersion = "14"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'REM 10/06/02 - 14 to 15
    If sBuildSubVersion = "14" Then
        sBuildSubVersion = "15"
        Call UpGradeData2_2from14to15
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'MLM 25/06/02 - 15 to 18 (16 & 17 were used for patches - no db upgrade)
    If sBuildSubVersion = "15" Then
        sBuildSubVersion = "18"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'MLM 28/06/02 - 18 to 19
    If sBuildSubVersion = "18" Then
        sBuildSubVersion = "19"
        Call UpgradeData2_2from18to19
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpgradeDataDatabase2_2", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpgradeDataDatabase2_1()
'---------------------------------------------------------------------
' Upgrade a 2.1 MACRO.mdb, checking first for 2.0 -> 2.1
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim sSQL As String
Dim sMacroVersion As String
Dim sBuildSubVersion As String
Dim sMsg As String

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM MACROControl"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        MsgBox ("Your Macro database is not valid. Macro is being closed down.")
        ExitMACRO
        MACROEnd
    End If
    
    sBuildSubVersion = rsTemp![BuildSubVersion]
    sMacroVersion = rsTemp![MACROVersion]
    rsTemp.Close
    Set rsTemp = Nothing

    'don't do anything if futre version
    If Val(sMacroVersion) >= 2.1 Then
        Exit Sub
    End If
    ' Check for version 2.0
    If sMacroVersion = "2.0" Then
        'REM 31/07/03 - No longer want this message to appear
        'sMsg = "You are about to upgrade your MACRO database from 2.0 to 2.1. Do you wish to continue?"
        'Select Case MsgBox(sMsg, vbQuestion + vbYesNo, gsDIALOG_TITLE)
            'Case vbYes
                Call UpGradeData2_0to2_1_1
                sMacroVersion = "2.1"
                sBuildSubVersion = "1"
            'Case vbNo
                'Call ExitMACRO
                'Call MACROEnd
        'End Select
    End If
    
    If sMacroVersion <> "2.1" And sMacroVersion <> "2.2" Then
        MsgBox ("Your Macro database is not valid. Macro is being closed down.")
        ExitMACRO
        MACROEnd
    End If
    
    'ADD NEW 2.1 UPGRADES HERE
    'Check for "01" as well as "1" only necessary for Sub Version 1
    ' NCj 5/10/00 Go straight from 1 to 3 (version 2 is defunct)
    If sBuildSubVersion = "1" Or sBuildSubVersion = "01" Then
        UpGradeData2_1from1to4
        sBuildSubVersion = "4"
    End If
              
    ' NCJ 5/10/00 - Version 2 must be thrown away
    ' TA 12/10/00 - Version 3 must be thrown away aswell
    If sBuildSubVersion = "2" Or sBuildSubVersion = "3" Then
        sMsg = "Your MACRO database is version 2.1." & sBuildSubVersion & ", which is no longer supported. "
        sMsg = sMsg & vbCrLf & "We cannot upgrade this database, so you'll have to throw it away."
        sMsg = sMsg & vbCrLf & "Sorry about that!"
        Call MsgBox(sMsg, vbApplicationModal, gsDIALOG_TITLE)
        ' Throw them out unless it's System Management (where DB version doesn't matter)
        If App.Title <> "MACRO_SM" Then
            Call ExitMACRO
            Call MACROEnd
        End If
    End If

    If sBuildSubVersion = "4" Or sBuildSubVersion = "5" Then
        UpGradeData2_1from4to6
        sBuildSubVersion = "6"
    End If
   
   ' NCJ 1/11/00 - Upgrade from 6 to 11
   ' NCJ 28/11/00 - Include 7,8,9,10 as well!
    Select Case sBuildSubVersion
    Case "6", "7", "8", "9", "10"
        sBuildSubVersion = "11"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    Case Else
        ' Not relevant here
    End Select
    
   ' NCJ 8/11/00 - Upgrade from 11 to 12
    If sBuildSubVersion = "11" Then
        Call UpGradeData2_1from11to12
        sBuildSubVersion = "12"
    End If
    
    'Mo Morris 20/11/00 - Upgrade from 12 to 13
    If sBuildSubVersion = "12" Then
        Call UpGradeData2_1from12to13
        sBuildSubVersion = "13"
    End If
    
    'Mo Morris 19/12/00 - Upgrade from 13,14,15,16 to 17
    Select Case sBuildSubVersion
    Case sBuildSubVersion = "13", "14", "15", "16"
        sBuildSubVersion = "17"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    Case Else
        ' Not relevant here
    End Select
   
    'TA 30/1/2001: Upgrade from 17 to 30 ' leaving room for Ronald builds
    If Val(sBuildSubVersion) >= 17 And Val(sBuildSubVersion) < 30 Then
        Call UpGradeData2_1from17to30
        sBuildSubVersion = "30"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    ' NCJ 8 Feb 01 - Upgrade to 32
    ' NCJ 19 Feb 01 - Upgrade to 34
    If Val(sBuildSubVersion) >= 30 And Val(sBuildSubVersion) <= 33 Then
        sBuildSubVersion = "34"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
   
   'TA 01/03/2001
    If sBuildSubVersion = "34" Then
        Call UpGradeData2_1from34to35
        sBuildSubVersion = "35"
    End If
   
   ' TA 10/03/2001 - version 35-36
   ' NCJ 23/3/01 - version 36-37
   ' NCJ 2/4/01 - version 38
   ' NCJ 10/4/01 - version 39
   ' TA 18/04/2001 - version 40
    If Val(sBuildSubVersion) >= 35 And Val(sBuildSubVersion) <= 39 Then
        sBuildSubVersion = "40"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

   
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpgradeDataDatabase2_1", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeData2_0to2_1_1()
'---------------------------------------------------------------------
' Upgrade 2.0.x (latest) MACRO.mdb to 2.1.1
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
  
    'Upgrade MacroVersion to [2.1]
    sSQL = "UPDATE MACROControl Set MacroVersion = '2.1'"
    MacroADODBConnection.Execute sSQL
    
    'Upgrade BuildSubVersion to [1]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '1'"
    MacroADODBConnection.Execute sSQL
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeData2_0to2_1_1", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpgradeDataDatabase2_0()
'---------------------------------------------------------------------
' Upgrade a 2.0 database to the last 2.0 database
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim sSQL As String
Dim sMacroVersion As String
Dim sBuildSubVersion As String

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM MACROControl"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        MsgBox ("Your Macro database is not valid. Macro is being closed down.")
        ExitMACRO
        MACROEnd
    End If
    
    sMacroVersion = rsTemp![MACROVersion]
    sBuildSubVersion = rsTemp![BuildSubVersion]
    rsTemp.Close
    Set rsTemp = Nothing
    
    If sMacroVersion <> "2.0" Then
        ' Not a 2.0 database
        Exit Sub
    End If
    
    'Check for need to upgrade from BuildSubVersion [15] to [16]
    If sBuildSubVersion = "15" Then
        sBuildSubVersion = "16"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'Check for need to upgrade from BuildSubVersion [16] to [17]
    If sBuildSubVersion = "16" Then
        sBuildSubVersion = "17"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'Check for need to upgrade from BuildSubVersion [17] to [18]
    If sBuildSubVersion = "17" Then
        UpGradeData17to18       ' This involves database changes
        sBuildSubVersion = "18"
    End If
    
    'Check for need to upgrade from BuildSubVersion [18] to [19]
    If sBuildSubVersion = "18" Then
        sBuildSubVersion = "19"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'Check for need to upgrade from BuildSubVersion [19] to [20]
    If sBuildSubVersion = "19" Then
        sBuildSubVersion = "20"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'Check for need to upgrade from BuildSubVersion [20] to [21]
    If sBuildSubVersion = "20" Then
        sBuildSubVersion = "21"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'Check for need to upgrade from BuildSubVersion [21] to [22]
    ' NCJ 8/3/00
    If sBuildSubVersion = "21" Then
        UpGradeData21to22       ' This involves database changes
        sBuildSubVersion = "22"
    End If

    'Check for need to upgrade from BuildSubVersion [22] to [23]
    ' WillC 14/3/00
    If sBuildSubVersion = "22" Then
        UpGradeData22to23       ' This involves database changes
        sBuildSubVersion = "23"
    End If

    'Check for need to upgrade from BuildSubVersion [23] to [27]
    ' NCJ 3/4/00
    If sBuildSubVersion = "23" Then
        sBuildSubVersion = "27"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    'Check for need to upgrade from BuildSubVersion [27] to [29]
    ' WillC 26/4/00
    If sBuildSubVersion = "27" Then
        UpGradeData27to29       ' This involves database changes
        sBuildSubVersion = "29"
    End If

    ' NCJ
    If sBuildSubVersion = "29" Then
        sBuildSubVersion = "31"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    ' NCJ
    If sBuildSubVersion = "31" Then
        sBuildSubVersion = "32"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'WillC 16/5/00
    'Check for need to upgrade from BuildSubVersion [32] to [33]
    If sBuildSubVersion = "32" Then
        UpGradeData32to33       ' This involves database changes
        sBuildSubVersion = "33"
    End If
    
    'WillC 26/5/00
    'Check for need to upgrade from BuildSubVersion [33] to [34]
    If sBuildSubVersion = "33" Then
        sBuildSubVersion = "34"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'WillC 30/5/00
    'Check for need to upgrade from BuildSubVersion [34] to [35]
    If sBuildSubVersion = "34" Then
        sBuildSubVersion = "35"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

    'NCJ 2/6/00
    If sBuildSubVersion = "35" Then
        sBuildSubVersion = "36"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
        'WillC 9/6/00
    If sBuildSubVersion = "36" Then
        sBuildSubVersion = "37"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

        'Nicky 15/6/00
    If sBuildSubVersion = "37" Then
        sBuildSubVersion = "39"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If

        'Nicky 16/6/00
    If sBuildSubVersion = "39" Then
        sBuildSubVersion = "40"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpgradeDataDatabase2_0", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeSecurity15to16()
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler

    'check for need to remove a trailing space in the RoleFunction table
    sSQL = "SELECT * FROM RoleFunction" _
        & " WHERE RoleCode = 'MacroUser'" _
        & " AND FunctionCode = 'F3002 '"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    If rsTemp.RecordCount = 1 Then
        rsTemp!FunctionCode = "F3002"
        rsTemp.Update
        'Debug.Print "F3002 changed"
    End If
    rsTemp.Close
    Set rsTemp = Nothing
    
    'Upgrade BuildSubVersion from [15] to [16]
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '16'"
    SecurityADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeSecurity15to16", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeData17to18()
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    'Correct an entry in table StandardDataFormat
    sSQL = "UPDATE StandardDataFormat Set DataTypeId = 0 " _
       & " WHERE StandardDataFormatId = 2"
    MacroADODBConnection.Execute sSQL
    
    'Change an entry in table StandardDataFormat
    sSQL = "UPDATE StandardDataFormat Set DataFormat = 'mm/dd/yyyy'" _
       & " WHERE StandardDataFormatId = 11"
    MacroADODBConnection.Execute sSQL
    
    'Add ImportTimeStamp to the 5 Patient data files
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'add to TrialSubject
        sSQL = "ALTER Table TrialSubject ADD COLUMN ImportTimeStamp DOUBLE"
        MacroADODBConnection.Execute sSQL
        'add to VisitInstance
        sSQL = "ALTER Table VisitInstance ADD COLUMN ImportTimeStamp DOUBLE"
        MacroADODBConnection.Execute sSQL
        'add to CRFPageInstance
        sSQL = "ALTER Table CRFPageInstance ADD COLUMN ImportTimeStamp DOUBLE"
        MacroADODBConnection.Execute sSQL
        'add to DataItemResponse
        sSQL = "ALTER Table DataItemResponse ADD COLUMN ImportTimeStamp DOUBLE"
        MacroADODBConnection.Execute sSQL
        'add to DataItemResponseHistory
        sSQL = "ALTER Table DataItemResponseHistory ADD COLUMN ImportTimeStamp DOUBLE"
        MacroADODBConnection.Execute sSQL
    Else    'SQLServer or Oracle
        'add to TrialSubject
        sSQL = "ALTER Table TrialSubject ADD ImportTimeStamp DECIMAL(16,10) NULL"
        MacroADODBConnection.Execute sSQL
        'add to VisitInstance
        sSQL = "ALTER Table VisitInstance ADD ImportTimeStamp DECIMAL(16,10) NULL"
        MacroADODBConnection.Execute sSQL
        'add to CRFPageInstance
        sSQL = "ALTER Table CRFPageInstance ADD ImportTimeStamp DECIMAL(16,10) NULL"
        MacroADODBConnection.Execute sSQL
        'add to DataItemResponse
        sSQL = "ALTER Table DataItemResponse ADD ImportTimeStamp DECIMAL(16,10) NULL"
        MacroADODBConnection.Execute sSQL
        'add to DataItemResponseHistory
        sSQL = "ALTER Table DataItemResponseHistory ADD ImportTimeStamp DECIMAL(16,10) NULL"
        MacroADODBConnection.Execute sSQL
    End If
  
    'Upgrade BuildSubVersion from [17] to [18]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '18'"
    MacroADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeData17to18", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeSecurity19to20()
'---------------------------------------------------------------------
' NCJ 16/2/00 Ignore errors in this routine because we unfortunately
' got an incompatibility between upgraded 19 DBs and script-generated 19 DBs
' (this change should have been in UpGradeSecurity18to19 but we missed the boat)
'---------------------------------------------------------------------
Dim sSQL As String

    On Error Resume Next
    
    sSQL = "Insert INTO Function (FunctionCode,Function) " _
        & " VALUES ('F2002','Disable user')"
    SecurityADODBConnection.Execute sSQL, dbFailOnError
    
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) " _
        & " VALUES ('MacroUser','F2002')"
    SecurityADODBConnection.Execute sSQL, dbFailOnError

    On Error GoTo ErrHandler
    
    'Upgrade BuildSubVersion from [19] to [20]
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '20'"
    SecurityADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeSecurity19to20", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeSecurity27to29()
'---------------------------------------------------------------------
'WillC 26/4/00
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    'Change the Function descriptions replace the word trial with study
    sSQL = " UPDATE Function SET Function = 'Assign user to study' WHERE FunctionCode = 'F2006'"
    SecurityADODBConnection.Execute sSQL, dbFailOnError
    sSQL = " UPDATE Function SET Function = 'Add site to study or study to site' WHERE FunctionCode = 'F4002'"
    SecurityADODBConnection.Execute sSQL, dbFailOnError
    sSQL = " UPDATE Function SET Function = 'Remove site from study' WHERE FunctionCode = 'F4003'"
    SecurityADODBConnection.Execute sSQL, dbFailOnError
    sSQL = " UPDATE Function SET Function = 'Change study status' WHERE FunctionCode = 'F4005'"
    SecurityADODBConnection.Execute sSQL, dbFailOnError
    
    'Add 4 new functions to the function table
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5013','View question audit trail')"
    SecurityADODBConnection.Execute sSQL, dbFailOnError
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5014','Overrule discrepancies')"
    SecurityADODBConnection.Execute sSQL, dbFailOnError
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5015','Add data entry question comment')"
    SecurityADODBConnection.Execute sSQL, dbFailOnError
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5016','View data entry question comments')"
    SecurityADODBConnection.Execute sSQL, dbFailOnError
    
    
    ' Add the 4 new functions to the MacroUser role in the RoleFunction table.
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5013')"
    SecurityADODBConnection.Execute sSQL, dbFailOnError
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5014')"
    SecurityADODBConnection.Execute sSQL, dbFailOnError
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5015')"
    SecurityADODBConnection.Execute sSQL, dbFailOnError
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5016')"
    SecurityADODBConnection.Execute sSQL, dbFailOnError
    
           
    'Upgrade BuildSubVersion from [27] to [29]
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '29'"
    SecurityADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeSecurity27to29", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeSecurity32to33()
'---------------------------------------------------------------------
' WillC 16/5/00
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    'Upgrade BuildSubVersion from [32] to [33]
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '33'"
    SecurityADODBConnection.Execute sSQL

    'Update the function text
    sSQL = "UPDATE Function Set Function = 'Access Data Review' Where FunctionCode = 'F1006'"
    SecurityADODBConnection.Execute sSQL
    
   ' WillC 18/5/00 New Functions
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5017','Create Discrepancy')"
    SecurityADODBConnection.Execute sSQL
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5018','Create SDV Remark')"
    SecurityADODBConnection.Execute sSQL
    
    
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5017')"
    SecurityADODBConnection.Execute sSQL
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5018')"
    SecurityADODBConnection.Execute sSQL

    

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurity32to33", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub


'---------------------------------------------------------------------
Private Sub UpGradeSecurity33to34()
'---------------------------------------------------------------------
' WillC 26/5/00
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsRoleFunction As ADODB.Recordset

    On Error GoTo ErrHandler
    
    'Upgrade BuildSubVersion from [33] to [34]
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '34'"
    SecurityADODBConnection.Execute sSQL

    'Update the function text
    sSQL = "UPDATE Function Set Function = 'System integrity check' Where FunctionCode = 'F5010'"
    SecurityADODBConnection.Execute sSQL
    
    sSQL = "UPDATE Function Set Function = 'Audit trail integrity check' Where FunctionCode = 'F5012'"
    SecurityADODBConnection.Execute sSQL
   
   'Insert New Function
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F3023','Update Arezzo from GGB import')"
    SecurityADODBConnection.Execute sSQL
        
        
    'Delete from RoleFunction
    sSQL = "DELETE * FROM RoleFunction WHERE FunctionCode = 'F5011'"
    SecurityADODBConnection.Execute sSQL
    
    'Open the RoleFunction table and use the recordset.update method to clear the deletion
    'out of the table so we can do delete to the function table without error further  below.
    sSQL = "SELECT * From RoleFunction "
    Set rsRoleFunction = New ADODB.Recordset
    rsRoleFunction.Open sSQL, SecurityADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    With rsRoleFunction
        .MoveFirst
        .MoveLast
        .Update
     If .RecordCount Then
     End If
        .Close
    End With
    Set rsRoleFunction = Nothing
        
    'Delete from Function
    sSQL = "DELETE  FROM Function WHERE FunctionCode = 'F5011'"
    SecurityADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurity33to34", "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeSecurity34to35()
'---------------------------------------------------------------------
' WIllC 30 May 00
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "UPDATE Function Set Function = 'Create SDV Mark' Where FunctionCode = 'F5018'"
    SecurityADODBConnection.Execute sSQL

    'Upgrade BuildSubVersion from [34] to [35]
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '35'"
    SecurityADODBConnection.Execute sSQL

    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurity34to35", "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeSecurity36to37()
'---------------------------------------------------------------------
' WIllC 9 June 00
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

   'Insert New Function
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5019','Use WORD Templates')"
    SecurityADODBConnection.Execute sSQL

    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5019')"
    SecurityADODBConnection.Execute sSQL

    'Upgrade BuildSubVersion from [36] to [37]
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '37'"
    SecurityADODBConnection.Execute sSQL

    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurity36to37", "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeSecurity2_0to2_1_1()
'---------------------------------------------------------------------
' WillC 1/8/00 SR3728
' This upgrades a 2.0.x (latest) Security database to 2.1.1
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    'Upgrade MACROVersion from [2.0] to [2.1]
    sSQL = "UPDATE SecurityControl Set MACROVersion = '2.1'"
    SecurityADODBConnection.Execute sSQL

    'Upgrade BuildSubVersion to 1
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '1'"
    SecurityADODBConnection.Execute sSQL
    
    ' Add new permissions
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F2009','Change System Properties')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F2010','View System Log')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F2011','Reset Password')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5020','View Subject Data')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F2012','View Site/Server Communication')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F2013','Restore Database')"
    SecurityADODBConnection.Execute (sSQL)

    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F2009')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F2010')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F2011')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5020')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F2012')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F2013')"
    SecurityADODBConnection.Execute (sSQL)

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurity2_0to2_1_1", "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeData21to22()
'---------------------------------------------------------------------

Dim sSQL As String

    On Error GoTo ErrHandler
    
    'drop the no duplcates index on TrialName in table TrialType
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        sSQL = "DROP INDEX TrialTypeName on TrialType"
    Else
        sSQL = "ALTER TABLE TrialType DROP CONSTRAINT TrialTypeName"
    End If
    MacroADODBConnection.Execute sSQL
    
    'drop the no duplicates index on PhaseName in table TrialPhase
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        sSQL = "DROP INDEX idx_PhaseName on TrialPhase"
    Else
        sSQL = "ALTER TABLE TrialPhase DROP CONSTRAINT PhaseName"
    End If
    MacroADODBConnection.Execute sSQL
  
    'Upgrade BuildSubVersion from [21] to [22]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '22'"
    MacroADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeData21to22", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeData22to23()
'---------------------------------------------------------------------
'WillC 14/3/00
'---------------------------------------------------------------------

Dim sSQL As String

    On Error GoTo ErrHandler
    
    ' Add the column to the StudyVisit table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        sSQL = "ALTER TABLE  StudyVisit  ADD Column VisitDatePrompt BYTE "
    ElseIf goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Then
        sSQL = sSQL & " ALTER TABLE  StudyVisit  ADD  VisitDatePrompt TINYINT  NULL DEFAULT 0"  'Alter Table can only allow nullable columns to be added
    ElseIf goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
        ' NOTE that you dont need to put COLUMN in the PLSQL statement
        sSQL = sSQL & " ALTER TABLE  StudyVisit  ADD (VisitDatePrompt NUMBER(1) DEFAULT 0)"
    End If

    MacroADODBConnection.Execute sSQL
    
  
    'Upgrade BuildSubVersion from [22] to [23]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '23'"
    MacroADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeData22to23", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeData27to29()
'---------------------------------------------------------------------
'WillC 26/4/00
'---------------------------------------------------------------------
Dim sSQL As String
Dim oTableDef As TableDef
Dim oMacroDatabase As Database

  If goUser.Database.DatabaseType = MACRODatabaseType.Access Then      ',NonExcusive,ReadWrite,Password
     Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(goUser.Database.DatabaseLocation, False, False, "MS Access;PWD=" & goUser.Database.DatabasePassword)
  End If

    ' Add the columns to the DataItemResponse table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        sSQL = "ALTER TABLE  DataItemResponse  ADD Column ValidationId SMALLINT  "   ' Cant Add a default with the Jet ALTER TABLE Statement
    ElseIf goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Then
        sSQL = "ALTER TABLE  DataItemResponse  ADD  ValidationId INTEGER  NULL DEFAULT 0"   'Alter Table can only allow nullable columns to be added
    ElseIf goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
        ' NOTE that you dont need to put COLUMN in the PLSQL statement
        sSQL = "ALTER TABLE  DataItemResponse  ADD (ValidationId NUMBER(8) DEFAULT 0)"
    End If

    MacroADODBConnection.Execute sSQL

    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        sSQL = "ALTER TABLE  DataItemResponse  ADD Column ValidationMessage MEMO "
    ElseIf goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Then
        sSQL = "ALTER TABLE  DataItemResponse  ADD  ValidationMessage TEXT "  'Alter Table can only allow nullable columns to be added
    ElseIf goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
        ' NOTE that you dont need to put COLUMN in the PLSQL statement
        'WillC 29/4/2000 changed ValidationMessage to VARCHAR(2000)
        sSQL = "ALTER TABLE  DataItemResponse  ADD (ValidationMessage VARCHAR(2000))"
    End If

    MacroADODBConnection.Execute sSQL
    
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        sSQL = "ALTER TABLE  DataItemResponse  ADD Column OverruleReason TEXT(255) "
    ElseIf goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Then
        sSQL = "ALTER TABLE  DataItemResponse  ADD  OverruleReason VARCHAR(255) "  'Alter Table can only allow nullable columns to be added
    ElseIf goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
        ' NOTE that you dont need to put COLUMN in the PLSQL statement
        sSQL = "ALTER TABLE  DataItemResponse  ADD (OverruleReason VARCHAR(255))"
    End If
    
    MacroADODBConnection.Execute sSQL
        
    
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
    
        ' TA 28/04/2000: code to get round refresh problem
        Set oMacroDatabase = Nothing
        Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(goUser.Database.DatabaseLocation, False, False, "MS Access;PWD=" & goUser.Database.DatabasePassword)
    
        oMacroDatabase.TableDefs.Refresh
        Set oTableDef = oMacroDatabase.TableDefs("DataItemResponse")
        With oTableDef
              .Fields("ValidationId").DefaultValue = 0
              .Fields("ValidationId").Required = False
              .Fields("ValidationMessage").AllowZeroLength = True
              .Fields("ValidationMessage").Required = False
              .Fields("OverruleReason").AllowZeroLength = True
              .Fields("OverruleReason").Required = False
        End With
    End If
    
    
'------ ' Add the columns to the DataItemResponseHistory table----------------------------
    
    
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        sSQL = "ALTER TABLE  DataItemResponseHistory  ADD Column ValidationId SMALLINT  "
    ElseIf goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Then
        sSQL = "ALTER TABLE  DataItemResponseHistory  ADD  ValidationId INTEGER  NULL DEFAULT 0"  'Alter Table can only allow nullable columns to be added
    ElseIf goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
        ' NOTE that you dont need to put COLUMN in the PLSQL statement
        sSQL = "ALTER TABLE  DataItemResponseHistory  ADD (ValidationId NUMBER(8) DEFAULT 0)"
    End If

    MacroADODBConnection.Execute sSQL

    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        sSQL = "ALTER TABLE  DataItemResponseHistory  ADD Column ValidationMessage MEMO "
    ElseIf goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Then
        sSQL = "ALTER TABLE  DataItemResponseHistory  ADD  ValidationMessage TEXT "  'Alter Table can only allow nullable columns to be added
    ElseIf goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
        ' NOTE that you dont need to put COLUMN in the PLSQL statement
        'WillC 29/4/2000 changed ValidationMessage to VARCHAR(2000)
        sSQL = "ALTER TABLE  DataItemResponseHistory  ADD (ValidationMessage VARCHAR(2000))"
    End If

    MacroADODBConnection.Execute sSQL
    
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        sSQL = "ALTER TABLE  DataItemResponseHistory  ADD Column OverruleReason TEXT(255) "
    ElseIf goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Then
        sSQL = "ALTER TABLE  DataItemResponseHistory  ADD  OverruleReason VARCHAR(255) " 'Alter Table can only allow nullable columns to be added
    ElseIf goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
        ' NOTE that you dont need to put COLUMN in the PLSQL statement
        sSQL = "ALTER TABLE  DataItemResponseHistory  ADD (OverruleReason VARCHAR(255))"
    End If
    
    MacroADODBConnection.Execute sSQL
    
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
    
        ' TA 28/04/2000: code to get round refresh problem
        Set oMacroDatabase = Nothing
        Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(goUser.Database.DatabaseLocation, False, False, "MS Access;PWD=" & goUser.Database.DatabasePassword)
    
        oMacroDatabase.TableDefs.Refresh
        Set oTableDef = oMacroDatabase.TableDefs("DataItemResponseHistory")
        With oTableDef
              .Fields("ValidationId").DefaultValue = 0
              .Fields("ValidationId").Required = False
              .Fields("ValidationMessage").AllowZeroLength = True
              .Fields("ValidationMessage").Required = False
              .Fields("OverruleReason").AllowZeroLength = True
              .Fields("OverruleReason").Required = False
        End With
    End If
    Set oMacroDatabase = Nothing
    

    'Upgrade BuildSubVersion from [27] to [29]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '29'"
    MacroADODBConnection.Execute sSQL
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                    "UpGradeData27to29", "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select


End Sub

'---------------------------------------------------------------------
Private Sub UpGradeData32to33()
'---------------------------------------------------------------------
' WillC 16/5/00
' If Any changes to the MACRO database make sure they are reflected
' in the RestoreSite module to keep it up to date
'---------------------------------------------------------------------
Dim sSQL As String
Dim oTableDef As TableDef
Dim oMacroDatabase As Database

    On Error GoTo ErrHandler
  
    'REM 07/06/02 - Leaving this as DAO as we never insert values into this table, only users do
    'Create the new Table MIMessage.
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        Set oMacroDatabase = Nothing
        Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(goUser.Database.DatabaseLocation, False, False, "MS Access;PWD=" & goUser.Database.DatabasePassword)
        sSQL = "CREATE TABLE MIMessage(MIMessageID INTEGER ,"
        sSQL = sSQL & " MIMessageSite TEXT(8),MIMessageSource SMALLINT,"
        sSQL = sSQL & " MIMessageType SMALLINT, MIMessageScope SMALLINT,"
        sSQL = sSQL & " MIMessageObjectID INTEGER, MIMessageObjectSource SMALLINT, MIMessagePriority SMALLINT,"
        sSQL = sSQL & " MIMessageTrialName TEXT(15), MIMessagePersonId INTEGER,"
        sSQL = sSQL & " MIMessageVisitId INTEGER, MIMessageVisitCycle SMALLINT,"
        sSQL = sSQL & " MIMessageCRFPageTaskID INTEGER, MIMessageResponseTaskId INTEGER,"
        sSQL = sSQL & " MIMessageResponseValue TEXT(255),MIMessageOCDiscrepancyID INTEGER, MIMessageCreated DOUBLE,"
        sSQL = sSQL & " MIMessageSent DOUBLE, MIMessageReceived DOUBLE,"
        sSQL = sSQL & " MIMessageHistory SMALLINT, MIMessageProcessed SMALLINT,"
        sSQL = sSQL & " MIMessageStatus SMALLINT, MIMessageText TEXT(255),"
        sSQL = sSQL & " MIMessageUserCode TEXT(50), MIMessageUserName TEXT(255),"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (MIMessageID,MIMessageSite,MIMessageSource))"
        oMacroDatabase.Execute sSQL, dbFailOnError
        
        oMacroDatabase.TableDefs.Refresh
        Set oTableDef = oMacroDatabase.TableDefs("MIMessage")
        With oTableDef
            .Fields("MIMessageID").DefaultValue = 0
            .Fields("MIMessageID").Required = True
            .Fields("MIMessagesite").AllowZeroLength = False
            .Fields("MIMessagesite").Required = True
            .Fields("MIMessageSource").DefaultValue = 0
            .Fields("MIMessageSource").Required = True
            .Fields("MIMessageType").DefaultValue = 0
            .Fields("MIMessageType").Required = True
            .Fields("MIMessageScope").DefaultValue = 4
            .Fields("MIMessageScope").Required = True
            .Fields("MIMessageObjectID").DefaultValue = 0
            .Fields("MIMessageObjectID").Required = True
            .Fields("MIMessagePriority").DefaultValue = 5
            .Fields("MIMessagePriority").Required = False
            .Fields("MIMessageTrialName").AllowZeroLength = True
            .Fields("MIMessageTrialName").Required = True
            .Fields("MIMessagePersonId").DefaultValue = 0
            .Fields("MIMessagePersonId").Required = False
            .Fields("MIMessageVisitId").DefaultValue = 0
            .Fields("MIMessageVisitId").Required = False
            .Fields("MIMessageVisitCycle").DefaultValue = 0
            .Fields("MIMessageVisitCycle").Required = False
            .Fields("MIMessageCRFPageTaskID").DefaultValue = 0
            .Fields("MIMessageCRFPageTaskID").Required = False
            .Fields("MIMessageResponseTaskId").DefaultValue = 0
            .Fields("MIMessageResponseTaskId").Required = False
            .Fields("MIMessageResponseValue").AllowZeroLength = True
            .Fields("MIMessageResponseValue").Required = False
            .Fields("MIMessageOCDiscrepancyID").Required = False
            .Fields("MIMessageOCDiscrepancyID").DefaultValue = 0
            .Fields("MIMessageCreated").Required = True
            .Fields("MIMessageCreated").DefaultValue = 0
            .Fields("MIMessageSent").Required = True
            .Fields("MIMessageSent").DefaultValue = 0
            .Fields("MIMessageReceived").Required = True
            .Fields("MIMessageReceived").DefaultValue = 0
            .Fields("MIMessageHistory").Required = True
            .Fields("MIMessageHistory").DefaultValue = 0
            .Fields("MIMessageProcessed").Required = True
            .Fields("MIMessageProcessed").DefaultValue = 0
            .Fields("MIMessageStatus").Required = False
            .Fields("MIMessageStatus").DefaultValue = 0
            .Fields("MIMessageText").AllowZeroLength = True
            .Fields("MIMessageText").Required = False
            .Fields("MIMessageUserCode").AllowZeroLength = False
            .Fields("MIMessageUserCode").Required = True
            .Fields("MIMessageUserName").AllowZeroLength = True
            .Fields("MIMessageUserName").Required = False
        End With
        Set oTableDef = Nothing
        
        sSQL = "CREATE INDEX idx_MIMessageID "
        sSQL = sSQL & " ON MIMessage ( MIMessageID)"
        oMacroDatabase.Execute sSQL, dbFailOnError
        
        sSQL = " CREATE INDEX idx_MIMessageUserCode "
        sSQL = sSQL & " ON MIMessage ( MIMessageUserCode )"
        oMacroDatabase.Execute sSQL, dbFailOnError
          
        sSQL = " CREATE INDEX idx_MIMessageObjectID "
        sSQL = sSQL & " ON MIMessage ( MIMessageObjectID )"
        oMacroDatabase.Execute sSQL, dbFailOnError
        
        Set oMacroDatabase = Nothing
    ElseIf goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Or MACRODatabaseType.Oracle80 Then
        sSQL = "CREATE TABLE MIMessage(MIMessageID INTEGER DEFAULT 0 ,"
        sSQL = sSQL & " MIMessageSite VARCHAR(8) ,MIMessageSource SMALLINT DEFAULT 0 ,"
        sSQL = sSQL & " MIMessageType SMALLINT DEFAULT 0 NOT NULL, MIMessageScope SMALLINT DEFAULT 4 NOT NULL,"
        sSQL = sSQL & " MIMessageObjectID INTEGER DEFAULT 0 NOT NULL, MIMessageObjectSource SMALLINT DEFAULT 0, MIMessagePriority SMALLINT DEFAULT 5 ,"
        sSQL = sSQL & " MIMessageTrialName VARCHAR(15) NULL, MIMessagePersonId INTEGER DEFAULT 0 NULL,"
        sSQL = sSQL & " MIMessageVisitId INTEGER DEFAULT 0 NULL, MIMessageVisitCycle SMALLINT DEFAULT 0 NULL,"
        sSQL = sSQL & " MIMessageCRFPageTaskID INTEGER DEFAULT 0 NULL, MIMessageResponseTaskId INTEGER DEFAULT 0 NULL,"
        sSQL = sSQL & " MIMessageResponseValue VARCHAR(255) NULL,MIMessageOCDiscrepancyID INTEGER DEFAULT 0, MIMessageCreated DECIMAL(16,10) DEFAULT 0  NOT NULL,"
        sSQL = sSQL & " MIMessageSent DECIMAL(16,10) DEFAULT 0 NOT NULL, MIMessageReceived DECIMAL(16,10) DEFAULT 0 NOT NULL,"
        sSQL = sSQL & " MIMessageHistory SMALLINT DEFAULT 0 NOT NULL, MIMessageProcessed SMALLINT DEFAULT 0 NOT NULL,"
        sSQL = sSQL & " MIMessageStatus SMALLINT DEFAULT 0 NULL, MIMessageText VARCHAR(255) NULL,"
        sSQL = sSQL & " MIMessageUserCode VARCHAR(50) NOT NULL, MIMessageUserName VARCHAR(255) NULL,"
        sSQL = sSQL & " CONSTRAINT PKMIMessage PRIMARY KEY "
        sSQL = sSQL & " (MIMessageID,MIMessageSite,MIMessageSource))"
        MacroADODBConnection.Execute sSQL
                
        sSQL = "CREATE INDEX idx_MIMessageID "
        sSQL = sSQL & " ON MIMessage (MIMessageID)"
        MacroADODBConnection.Execute sSQL
        
        sSQL = " CREATE INDEX idx_MIMessageUserCode "
        sSQL = sSQL & " ON MIMessage ( MIMessageUserCode )"
        MacroADODBConnection.Execute sSQL
          
        sSQL = " CREATE INDEX idx_MIMessageObjectID "
        sSQL = sSQL & " ON MIMessage ( MIMessageObjectID )"
        MacroADODBConnection.Execute sSQL
    End If

    'Upgrade BuildSubVersion from [32] to [33]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '33'"
    MacroADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeData32to33", "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub


'---------------------------------------------------------------------
Private Sub UpGradeDataToSubVersion(sSubVersion As String)
'---------------------------------------------------------------------
' Nicky 1/9/00
' Generic routine to update the build subversion in MACROControl in MACRO.mdb
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
  
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '" & sSubVersion & "'"
    MacroADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeDataToSubVersion" & sSubVersion, "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeSecurityToSubVersion(sSubVersion As String)
'---------------------------------------------------------------------
' Nicky 1/9/00
' Generic routine to update the build subversion in SecurityControl in Security.mdb
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '" & sSubVersion & "'"
    SecurityADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurityToSubVersion" & sSubVersion, "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeData2_1from1to4()
'---------------------------------------------------------------------
' NCJ 5/10/00 Renamed from UpGradeData2_1from1to2 since version 2 no longer supported
' Removed long integer identifiers and use codes instead
' TA 12/10/2000 Renamed from UpGradeData2_1from1to3 since version 3 no longer supported
'---------------------------------------------------------------------
Dim sSQL As String
Dim oTableDef As TableDef
Dim oMacroDatabase As Database

    On Error GoTo ErrHandler
    
    'Remove the existing Laboratory table, that existed prior to version 2.1.2, but was not being used
    sSQL = "DROP Table Laboratory"
    MacroADODBConnection.Execute sSQL
    
    'Create the new Laboratory table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table Laboratory (LaboratoryCode TEXT(15),"
        sSQL = sSQL & " LaboratoryDescription TEXT(255),"
        'TA 11/10/2000: two new columns for laboratory
        sSQL = sSQL & " Site TEXT(8), Changed SMALLINT DEFAULT 0,"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (LaboratoryCode))"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError
'
'        oMacroDatabase.TableDefs.Refresh
'        With oMacroDatabase.TableDefs("Laboratory")
'            .Fields("Changed").DefaultValue = 0
'        End With
'
'        Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table Laboratory (LaboratoryCode VARCHAR(15),"
        sSQL = sSQL & " LaboratoryDescription VARCHAR(255),"
        'TA 11/10/2000: two new columns for laboratory
        sSQL = sSQL & " Site VARCHAR(8), Changed INTEGER DEFAULT 0,"
        sSQL = sSQL & " CONSTRAINT PKLaboratory PRIMARY KEY"
        sSQL = sSQL & " (LaboratoryCode))"
        MacroADODBConnection.Execute sSQL
    End If
    
    'Remove the existing NormalRange table, that existed prior to version 2.1.2, but was not being used
    sSQL = "DROP Table NormalRange"
    MacroADODBConnection.Execute sSQL
    
    'Create the new NormalRange table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table NormalRange (NormalRangeId INTEGER,"
        sSQL = sSQL & " LaboratoryCode TEXT(15) , ClinicalTestCode TEXT(15),"
        sSQL = sSQL & " NormalRangeGender SMALLINT,"
        sSQL = sSQL & " NormalRangeAgeMin SMALLINT, NormalRangeAgeMax SMALLINT,"
        sSQL = sSQL & " NormalRangeEffectiveStart DOUBLE, NormalRangeEffectiveEnd DOUBLE,"
        sSQL = sSQL & " NormalRangeNormalMin DOUBLE, NormalRangeNormalMax DOUBLE,"
        sSQL = sSQL & " NormalRangeFeasibleMin DOUBLE, NormalRangeFeasibleMax DOUBLE,"
        sSQL = sSQL & " NormalRangeAbsoluteMin DOUBLE, NormalRangeAbsoluteMax DOUBLE,"
        sSQL = sSQL & " NormalRangePercent SMALLINT,"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & "(NormalRangeId,LaboratoryCode))"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError
'
'        oMacroDatabase.TableDefs.Refresh
        
        sSQL = " CREATE INDEX idx_LaboratoryCode "
        sSQL = sSQL & "ON NormalRange ( LaboratoryCode )"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError

        sSQL = " CREATE INDEX idx_ClinicalTestCode "
        sSQL = sSQL & "ON NormalRange ( ClinicalTestCode )"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError
        
'        Set oMacroDatabase = Nothing
        
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table NormalRange (NormalRangeId INTEGER,"
        sSQL = sSQL & " LaboratoryCode VARCHAR(15), ClinicalTestCode VARCHAR(15),"
        sSQL = sSQL & " NormalRangeGender INTEGER,"
        sSQL = sSQL & " NormalRangeAgeMin INTEGER, NormalRangeAgeMax INTEGER,"
        sSQL = sSQL & " NormalRangeEffectiveStart DECIMAL(16,10), NormalRangeEffectiveEnd DECIMAL(16,10),"
        sSQL = sSQL & " NormalRangeNormalMin DECIMAL(16,10), NormalRangeNormalMax DECIMAL(16,10),"
        sSQL = sSQL & " NormalRangeFeasibleMin DECIMAL(16,10), NormalRangeFeasibleMax DECIMAL(16,10),"
        sSQL = sSQL & " NormalRangeAbsoluteMin DECIMAL(16,10), NormalRangeAbsoluteMax DECIMAL(16,10),"
        sSQL = sSQL & " NormalRangePercent INTEGER,"
        sSQL = sSQL & " CONSTRAINT PKNormalRange PRIMARY KEY "
        sSQL = sSQL & " (NormalRangeId, LaboratoryCode))"
        MacroADODBConnection.Execute sSQL
        
        sSQL = "CREATE INDEX idx_NR_LaboratoryCode "
        sSQL = sSQL & "ON NormalRange ( LaboratoryCode )"
        MacroADODBConnection.Execute sSQL

        sSQL = "CREATE INDEX idx_NR_ClinicalTestCode "
        sSQL = sSQL & "ON NormalRange ( ClinicalTestCode )"
        MacroADODBConnection.Execute sSQL
    End If
    
    
    'Create the new SiteLaboratory table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table SiteLaboratory (Site Text(8),"
        sSQL = sSQL & " LaboratoryCode Text(15),"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (Site,LaboratoryCode))"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError
'
'        oMacroDatabase.TableDefs.Refresh
'
'        Set oMacroDatabase = Nothing
        
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table SiteLaboratory (Site VARCHAR(8),"
        sSQL = sSQL & " LaboratoryCode VARCHAR(15),"
        sSQL = sSQL & " CONSTRAINT PKSiteLaboratory PRIMARY KEY "
        sSQL = sSQL & " (Site,LaboratoryCode))"
        MacroADODBConnection.Execute sSQL
    End If
    
    
    'Create the new ClinicalTestGroup table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table ClinicalTestGroup (ClinicalTestGroupCode TEXT(15),"
        sSQL = sSQL & " ClinicalTestGroupDescription TEXT(255),"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTestGroupCode))"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError
'
'        oMacroDatabase.TableDefs.Refresh
'
'        Set oMacroDatabase = Nothing
        
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table ClinicalTestGroup (ClinicalTestGroupCode VARCHAR(15),"
        sSQL = sSQL & " ClinicalTestGroupDescription VARCHAR(255),"
        sSQL = sSQL & " CONSTRAINT PKClinicalTestGroup PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTestGroupCode))"
        MacroADODBConnection.Execute sSQL
    End If
    
    'Create the new ClinicalTest table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table ClinicalTest (ClinicalTestCode TEXT(15),"
        sSQL = sSQL & " ClinicalTestDescription TEXT(255),"
        sSQL = sSQL & " ClinicalTestGroupCode TEXT(15),"
        sSQL = sSQL & " Unit TEXT(15),"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTestCode))"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError
'
'        oMacroDatabase.TableDefs.Refresh
    
        sSQL = " CREATE INDEX idx_ClinicalTestGroupCode "
        sSQL = sSQL & "ON ClinicalTest ( ClinicalTestGroupCode )"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError
'
'        Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table ClinicalTest (ClinicalTestCode VARCHAR(15),"
        sSQL = sSQL & " ClinicalTestDescription VARCHAR(255),"
        sSQL = sSQL & " ClinicalTestGroupCode VARCHAR(15),"
        sSQL = sSQL & " Unit VARCHAR(15),"
        sSQL = sSQL & " CONSTRAINT PKClinicalTest PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTestCode))"
        MacroADODBConnection.Execute sSQL
    
        sSQL = " CREATE INDEX idx_CT_ClinicalTestGroupCode "
        sSQL = sSQL & "ON ClinicalTest ( ClinicalTestGroupCode )"
        MacroADODBConnection.Execute sSQL
    End If
    
    'Create the new CTCScheme table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table CTCScheme (CTCSchemeCode TEXT(15),"
        sSQL = sSQL & " CTCSchemeDescription TEXT(255),"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (CTCSchemeCode))"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError
'
'        oMacroDatabase.TableDefs.Refresh
'
'        Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table CTCScheme (CTCSchemeCode VARCHAR(15),"
        sSQL = sSQL & " CTCSchemeDescription VARCHAR(255),"
        sSQL = sSQL & " CONSTRAINT PKCTCScheme PRIMARY KEY "
        sSQL = sSQL & " (CTCSchemeCode))"
        MacroADODBConnection.Execute sSQL
    End If
    
    
    'Create the new CTC table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table CTC (CTCId INTEGER ,"
        sSQL = sSQL & " CTCSchemeCode TEXT(15),"
        sSQL = sSQL & " ClinicalTestCode TEXT(15),"
        sSQL = sSQL & " CTCGrade SMALLINT,"
        sSQL = sSQL & " CTCMin DOUBLE,"
        sSQL = sSQL & " CTCMax DOUBLE,"
        sSQL = sSQL & " CTCMinType SMALLINT,"
        sSQL = sSQL & " CTCMaxType SMALLINT,"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (CTCId,CTCSchemeCode))"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
'        oMacroDatabase.Execute sSQL, dbFailOnError
'
'        oMacroDatabase.TableDefs.Refresh
    
        sSQL = " CREATE INDEX idx_CTCSchemeCode "
        sSQL = sSQL & "ON CTC ( CTCSchemeCode )"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        sSQL = " CREATE INDEX idx_ClinicalTestCode "
        sSQL = sSQL & "ON CTC ( ClinicalTestCode )"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
                
        'Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table CTC (CTCId INTEGER,"
        sSQL = sSQL & " CTCSchemeCode VARCHAR(15),"
        sSQL = sSQL & " ClinicalTestCode VARCHAR(15),"
        sSQL = sSQL & " CTCGrade INTEGER,"
        sSQL = sSQL & " CTCMin DECIMAL(16,10),"
        sSQL = sSQL & " CTCMax DECIMAL(16,10),"
        sSQL = sSQL & " CTCMinType INTEGER,"
        sSQL = sSQL & " CTCMaxType INTEGER,"
        sSQL = sSQL & " CONSTRAINT PKCTC PRIMARY KEY "
        sSQL = sSQL & " (CTCId,CTCSchemeCode))"
        MacroADODBConnection.Execute sSQL
    
        sSQL = " CREATE INDEX idx_CTC_CTCSchemeCode "
        sSQL = sSQL & "ON CTC ( CTCSchemeCode )"
        MacroADODBConnection.Execute sSQL
        
        sSQL = " CREATE INDEX idx_CTC_ClinicalTestCode "
        sSQL = sSQL & "ON CTC ( ClinicalTestCode )"
        MacroADODBConnection.Execute sSQL
    End If
    
    
    'Amend the StudyDefinition table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "ALTER Table StudyDefinition ADD COLUMN CTCSchemeCode TEXT(15)"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        sSQL = "ALTER Table StudyDefinition ADD COLUMN DOBExpr MEMO"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        sSQL = "ALTER Table StudyDefinition ADD COLUMN GenderExpr MEMO"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
        
        'Set oMacroDatabase = Nothing
        
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table StudyDefinition ADD CTCSchemeCode VARCHAR(15)"
        MacroADODBConnection.Execute sSQL
        'Oracle / SQL server specifics
        If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Then
            sSQL = "ALTER Table StudyDefinition ADD DOBExpr TEXT"
            MacroADODBConnection.Execute sSQL
            sSQL = "ALTER Table StudyDefinition ADD GenderExpr TEXT"
            MacroADODBConnection.Execute sSQL
        ElseIf goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
            sSQL = "ALTER Table StudyDefinition ADD DOBExpr VARCHAR(2000)"
            MacroADODBConnection.Execute sSQL
            sSQL = "ALTER Table StudyDefinition ADD GenderExpr VARCHAR(2000)"
            MacroADODBConnection.Execute sSQL
        End If
    
    End If
    
    
    'Amend the DataItem table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "ALTER Table DataItem ADD COLUMN ClinicalTestCode TEXT(15)"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
        
        'Set oMacroDatabase = Nothing
    
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table DataItem ADD ClinicalTestCode VARCHAR(15)"
        MacroADODBConnection.Execute sSQL
        
    End If
    
    
    'Amend the CRFElement table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using ADO (no need for DAO facilities)
        sSQL = "ALTER Table CRFElement ADD COLUMN ClinicalTestDateExpr Memo"
        MacroADODBConnection.Execute sSQL
    Else    'SQL or Oracle, using ADO
        'Oracle / SQL server specifics
        If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Then
            sSQL = "ALTER Table CRFElement ADD ClinicalTestDateExpr TEXT"
        ElseIf goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
            sSQL = "ALTER Table CRFElement ADD ClinicalTestDateExpr VARCHAR(2000)"
        End If
        MacroADODBConnection.Execute sSQL
    End If
    
    
    'Amend the CRFPageInstance table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "ALTER Table CRFPageInstance ADD COLUMN LaboratoryCode TEXT(15)"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
            
        'Set oMacroDatabase = Nothing
        
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table CRFPageInstance ADD LaboratoryCode VARCHAR(15)"
        MacroADODBConnection.Execute sSQL
            
    End If
    
    'Amend the TrialSubject table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        ' TA 12/10/2000: no longer deleting the current Gender and DateofBirth columns - they should not be used though
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "DROP INDEX idx_Gender on TrialSubject"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        sSQL = "Alter Table TrialSubject ADD COLUMN SubjectGender SMALLINT"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        'Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "DROP INDEX idx_Gender"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table TrialSubject ADD SubjectGender INTEGER"
        MacroADODBConnection.Execute sSQL
    End If
    
    
    'Amend the DataItemResponse table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using ADO (no need for DAO facilities)
        sSQL = "ALTER Table DataItemResponse ADD COLUMN LabResult TEXT(1)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table DataItemResponse ADD COLUMN CTCGrade SMALLINT"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table DataItemResponse ADD COLUMN ClinicalTestDate DOUBLE"
        MacroADODBConnection.Execute sSQL
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table DataItemResponse ADD LabResult VARCHAR(1)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table DataItemResponse ADD CTCGrade INTEGER"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table DataItemResponse ADD ClinicalTestDate DECIMAL(16,10)"
        MacroADODBConnection.Execute sSQL
    End If
    
    
    'Amend the DataItemResponseHistory table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using ADO (no need for DAO facilities)
        sSQL = "ALTER Table DataItemResponseHistory ADD COLUMN LabResult TEXT(1)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table DataItemResponseHistory ADD COLUMN CTCGrade SMALLINT"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table DataItemResponseHistory ADD COLUMN ClinicalTestDate DOUBLE"
        MacroADODBConnection.Execute sSQL
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table DataItemResponseHistory ADD LabResult VARCHAR(1)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table DataItemResponseHistory ADD CTCGrade INTEGER"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table DataItemResponseHistory ADD ClinicalTestDate DECIMAL(16,10)"
        MacroADODBConnection.Execute sSQL
    End If
    
    'TA 12/10/2000: Create the new NewDbColumn table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'REM 07/06/02 - Changed to ADO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        Call CreateNewDBColumn2_1_4(MacroADODBConnection, True)
        'Set oMacroDatabase = Nothing
        
    Else    'SQL or Oracle, using ADO
        Call CreateNewDBColumn2_1_4(MacroADODBConnection, False)
    End If

    
    'TA 12/10/2000: Create the new SiteUser table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table SiteUser (Site Text(8),"
        sSQL = sSQL & " UserCode Text(15),"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (Site,UserCode))"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
        
        'Set oMacroDatabase = Nothing
        
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table SiteUser (Site VARCHAR(8),"
        sSQL = sSQL & " UserCode VARCHAR(15),"
        sSQL = sSQL & " CONSTRAINT PKSiteUser PRIMARY KEY "
        sSQL = sSQL & " (Site,UserCode))"
        MacroADODBConnection.Execute sSQL
    End If
    
    'TA 12/10/2000: MACROTAble Table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'REM 07/06/02 - Changed to ADO
        Call CreateMacroTablev2_1_4(MacroADODBConnection, True)
    Else    'SQL or Oracle, using ADO
        Call CreateMacroTablev2_1_4(MacroADODBConnection, False)
    End If

    'TA 13/10/2000: drop old import/export tables
        If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(goUser.Database.DatabaseLocation, False, False, "MS Access;PWD=" & goUser.Database.DatabasePassword)

        oMacroDatabase.Execute "DROP TABLE SDDExportImport", dbFailOnError
        oMacroDatabase.Execute "DROP TABLE PRDExportImport", dbFailOnError
        
        oMacroDatabase.TableDefs.Refresh
        
        Set oMacroDatabase = Nothing
        
    Else    'SQL or Oracle, using ADO
        MacroADODBConnection.Execute "DROP TABLE SDDExportImport"
        MacroADODBConnection.Execute "DROP TABLE PRDExportImport"
    End If

    'Upgrade BuildSubVersion to [4]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '4'"
    MacroADODBConnection.Execute sSQL


Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeData2_1from1to4", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeData2_1from4to6()
'---------------------------------------------------------------------
' NCJ 25/10/00 Upgrade 2.1. from 4 to 6
' Add SegmentIds for Units and UnitConversionFactors
'---------------------------------------------------------------------
Dim sSQL As String

    sSQL = "UPDATE MACROTable Set SegmentId = '550', STYDEF = 1 WHERE TableName = 'Units'"
    MacroADODBConnection.Execute sSQL

    sSQL = "UPDATE MACROTable Set SegmentId = '560', STYDEF = 1 WHERE TableName = 'UnitConversionFactors'"
    MacroADODBConnection.Execute sSQL
    
    'Upgrade BuildSubVersion to [6]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '6'"
    MacroADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeData2_1from4to6", "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'---------------------------------------------------------------------
Private Sub UpGradeData2_1from11to12()
'---------------------------------------------------------------------
' NCJ 25/10/00 Upgrade 2.1. from 11 to 12
' Add ClinicalTest, ClinicalTestGroup and Units to LDD export
'---------------------------------------------------------------------
Dim sSQL As String

    sSQL = "UPDATE MACROTable Set LABDEF = 1 WHERE TableName = 'Units'"
    MacroADODBConnection.Execute sSQL

    sSQL = "UPDATE MACROTable Set LABDEF = 1 WHERE TableName = 'ClinicalTest'"
    MacroADODBConnection.Execute sSQL
    
    sSQL = "UPDATE MACROTable Set LABDEF = 1 WHERE TableName = 'ClinicalTestGroup'"
    MacroADODBConnection.Execute sSQL
    
    'Upgrade BuildSubVersion to [12]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '12'"
    MacroADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeData2_1from11to12", "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub


'---------------------------------------------------------------------
Private Sub UpGradeSecurity2_1from1to2()
'---------------------------------------------------------------------
Dim sSQL As String


    On Error GoTo ErrHandler
    
    ' Add new permissions
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F6001','Maintain Laboratories')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F6002','Maintain CTC Schemes')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F6003','Maintain Clinical Tests')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F6004','Maintain Normal Ranges')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F6005','Maintain Common Toxicity Criteria')"
    SecurityADODBConnection.Execute (sSQL)

    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F6001')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F6002')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F6003')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F6004')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F6005')"
    SecurityADODBConnection.Execute (sSQL)
    
    'Upgrade BuildSubVersion to 2
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '2'"
    SecurityADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurity2_1from1to2", "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeData2_1from12to13()
'---------------------------------------------------------------------
Dim sSQL As String
Dim oMacroDatabase As Database

    On Error GoTo ErrHandler
    
    'Create the new Eligibility table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table Eligibility ("
        sSQL = sSQL & " ClinicalTrialId INTEGER,"
        sSQL = sSQL & " VersionId SMALLINT,"
        sSQL = sSQL & " EligibilityCode TEXT(15),"
        sSQL = sSQL & " RandomisationCode TEXT(15),"
        sSQL = sSQL & " Flag SMALLINT,"
        sSQL = sSQL & " Condition MEMO,"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialId,VersionId,EligibilityCode))"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
        'Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table Eligibility ("
        sSQL = sSQL & " ClinicalTrialId INTEGER,"
        sSQL = sSQL & " VersionId INTEGER,"
        sSQL = sSQL & " EligibilityCode VARCHAR(15),"
        sSQL = sSQL & " RandomisationCode VARCHAR(15),"
        sSQL = sSQL & " Flag INTEGER,"
        If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Then
            sSQL = sSQL & " Condition  TEXT,"
        ElseIf goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
            sSQL = sSQL & " Condition  VARCHAR(2000),"
        End If
        sSQL = sSQL & " CONSTRAINT PKEligibility PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialId,VersionId,EligibilityCode))"
        MacroADODBConnection.Execute sSQL
    End If
    
    
    'Create the new UniquenessCheck table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table UniquenessCheck ("
        sSQL = sSQL & " ClinicalTrialId INTEGER,"
        sSQL = sSQL & " VersionId SMALLINT,"
        sSQL = sSQL & " CheckCode TEXT(15),"
        sSQL = sSQL & " Expression MEMO,"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialId,VersionId,CheckCode))"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
        'Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table UniquenessCheck ("
        sSQL = sSQL & " ClinicalTrialId INTEGER,"
        sSQL = sSQL & " VersionId INTEGER,"
        sSQL = sSQL & " CheckCode VARCHAR(15),"
        If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Then
            sSQL = sSQL & " Expression  TEXT,"
        ElseIf goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
            sSQL = sSQL & " Expression  VARCHAR(2000),"
        End If
        sSQL = sSQL & " CONSTRAINT PKUniquenessCheck PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialId,VersionId,CheckCode))"
        MacroADODBConnection.Execute sSQL
    End If
    
    
    'Create the new SubjectNumbering table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table SubjectNumbering ("
        sSQL = sSQL & " ClinicalTrialId INTEGER,"
        sSQL = sSQL & " VersionId SMALLINT,"
        sSQL = sSQL & " StartNumber INTEGER,"
        sSQL = sSQL & " NumberWidth SMALLINT,"
        sSQL = sSQL & " Prefix MEMO,"
        sSQL = sSQL & " UsePrefix SMALLINT,"
        sSQL = sSQL & " Suffix MEMO,"
        sSQL = sSQL & " UseSuffix SMALLINT,"
        sSQL = sSQL & " TriggerVisitId INTEGER,"
        sSQL = sSQL & " TriggerFormId INTEGER,"
        sSQL = sSQL & " UseRegistration SMALLINT,"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialId,VersionId))"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
        'Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table SubjectNumbering ("
        sSQL = sSQL & " ClinicalTrialId INTEGER,"
        sSQL = sSQL & " VersionId INTEGER,"
        sSQL = sSQL & " StartNumber INTEGER,"
        sSQL = sSQL & " NumberWidth INTEGER,"
        If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Then
            sSQL = sSQL & " Prefix  TEXT,"
        ElseIf goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
            sSQL = sSQL & " Prefix  VARCHAR(2000),"
        End If
        sSQL = sSQL & " UsePrefix INTEGER,"
        If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Then
            sSQL = sSQL & " Suffix  TEXT,"
        ElseIf goUser.Database.DatabaseType = MACRODatabaseType.Oracle80 Then
            sSQL = sSQL & " Suffix  VARCHAR(2000),"
        End If
        sSQL = sSQL & " UseSuffix INTEGER,"
        sSQL = sSQL & " TriggerVisitId INTEGER,"
        sSQL = sSQL & " TriggerFormId INTEGER,"
        sSQL = sSQL & " UseRegistration INTEGER,"
        sSQL = sSQL & " CONSTRAINT PKSubjectNumbering PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialId,VersionId))"
        MacroADODBConnection.Execute sSQL
    End If
    
    
    'Create the new RSSubjectIdentifier table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table RSSubjectIdentifier ("
        sSQL = sSQL & " ClinicalTrialName TEXT(15),"
        sSQL = sSQL & " TrialSite TEXT(8),"
        sSQL = sSQL & " PersonId INTEGER,"
        sSQL = sSQL & " SubjectIdentifier TEXT(255),"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialName,TrialSite,PersonId))"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
        'Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table RSSubjectIdentifier ("
        sSQL = sSQL & " ClinicalTrialName VARCHAR(15),"
        sSQL = sSQL & " TrialSite VARCHAR(8),"
        sSQL = sSQL & " PersonId INTEGER,"
        sSQL = sSQL & " SubjectIdentifier VARCHAR(255),"
        sSQL = sSQL & " CONSTRAINT PKRSSubjectIdentifier PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialName,TrialSite,PersonId))"
        MacroADODBConnection.Execute sSQL
    End If
    
    
    'Create the new RSNextNumber table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table RSNextNumber ("
        sSQL = sSQL & " ClinicalTrialName TEXT(15),"
        sSQL = sSQL & " NextNumberId INTEGER,"
        sSQL = sSQL & " Prefix TEXT(255),"
        sSQL = sSQL & " Suffix TEXT(255),"
        sSQL = sSQL & " NextNumber INTEGER,"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialName,NextNumberId))"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
        'Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table RSNextNumber ("
        sSQL = sSQL & " ClinicalTrialName VARCHAR(15),"
        sSQL = sSQL & " NextNumberId INTEGER,"
        sSQL = sSQL & " Prefix VARCHAR(255),"
        sSQL = sSQL & " Suffix VARCHAR(255),"
        sSQL = sSQL & " NextNumber INTEGER,"
        sSQL = sSQL & " CONSTRAINT PKRSNextNumber PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialName,NextNumberId))"
        MacroADODBConnection.Execute sSQL
    End If
    
    
    'Create the new RSUniquenessCheck table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        'Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(gUser.DatabasePath, False, False, "MS Access;PWD=" & gUser.DatabasePassword)
        sSQL = "CREATE Table RSUniquenessCheck ("
        sSQL = sSQL & " ClinicalTrialName TEXT(15),"
        sSQL = sSQL & " TrialSite TEXT(8),"
        sSQL = sSQL & " PersonId INTEGER,"
        sSQL = sSQL & " CheckCode TEXT(15),"
        sSQL = sSQL & " CheckValue TEXT(255),"
        sSQL = sSQL & " CheckTimeStamp DOUBLE,"
        sSQL = sSQL & " CONSTRAINT PrimaryKey PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialName,TrialSite,PersonId,CheckCode))"
        MacroADODBConnection.Execute sSQL 'REM 06/06/02 - Changed to ADO
        'oMacroDatabase.Execute sSQL, dbFailOnError
        
        'oMacroDatabase.TableDefs.Refresh
        'Set oMacroDatabase = Nothing
    Else    'SQL or Oracle, using ADO
        sSQL = "CREATE Table RSUniquenessCheck ("
        sSQL = sSQL & " ClinicalTrialName VARCHAR(15),"
        sSQL = sSQL & " TrialSite VARCHAR(8),"
        sSQL = sSQL & " PersonId INTEGER,"
        sSQL = sSQL & " CheckCode VARCHAR(15),"
        sSQL = sSQL & " CheckValue VARCHAR(255),"
        sSQL = sSQL & " CheckTimeStamp DECIMAL(16,10),"
        sSQL = sSQL & " CONSTRAINT PKRSUniquenessCheck PRIMARY KEY "
        sSQL = sSQL & " (ClinicalTrialName,TrialSite,PersonId,CheckCode))"
        MacroADODBConnection.Execute sSQL
    End If
    
    'Amend the TrialSubject table, add field RegistrationStatus
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using ADO (no need for DAO facilities)
        sSQL = "ALTER Table TrialSubject ADD COLUMN RegistrationStatus SMALLINT"
        MacroADODBConnection.Execute sSQL
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table TrialSubject ADD RegistrationStatus INTEGER"
        MacroADODBConnection.Execute sSQL
    End If
    
    
    'Amend the MIMessage table, add field MIMessageResponseTimeStamp
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using ADO (no need for DAO facilities)
        sSQL = "ALTER Table MIMessage ADD COLUMN MIMessageResponseTimeStamp DOUBLE"
        MacroADODBConnection.Execute sSQL
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table MIMessage ADD MIMessageResponseTimeStamp DECIMAL(16,10)"
        MacroADODBConnection.Execute sSQL
    End If
    
    
    'Amend the DataItemResponse table, add field LaboratoryCode
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using ADO (no need for DAO facilities)
        sSQL = "ALTER Table DataItemResponse ADD COLUMN LaboratoryCode TEXT(15)"
        MacroADODBConnection.Execute sSQL
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table DataItemResponse ADD LaboratoryCode VARCHAR(15)"
        MacroADODBConnection.Execute sSQL
    End If
    
    'Amend the DataItemResponseHistory table, add field LaboratoryCode
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using ADO (no need for DAO facilities)
        sSQL = "ALTER Table DataItemResponseHistory ADD COLUMN LaboratoryCode TEXT(15)"
        MacroADODBConnection.Execute sSQL
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table DataItemResponseHistory ADD LaboratoryCode VARCHAR(15)"
        MacroADODBConnection.Execute sSQL
    End If
    
    'Amend the StudyDfinition table
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using ADO (no need for DAO facilities)
        sSQL = "ALTER Table StudyDefinition ADD COLUMN RRServerType SMALLINT"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table StudyDefinition ADD COLUMN RRHTTPAddress TEXT(255)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table StudyDefinition ADD COLUMN RRUserName TEXT(50)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table StudyDefinition ADD COLUMN RRPassword TEXT(50)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table StudyDefinition ADD COLUMN RRProxyServer TEXT(255)"
        MacroADODBConnection.Execute sSQL
    Else    'SQL or Oracle, using ADO
        sSQL = "ALTER Table StudyDefinition ADD RRServerType INTEGER"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table StudyDefinition ADD RRHTTPAddress VARCHAR(255)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table StudyDefinition ADD RRUserName VARCHAR(50)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table StudyDefinition ADD RRPassword VARCHAR(50)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table StudyDefinition ADD RRProxyServer VARCHAR(255)"
        MacroADODBConnection.Execute sSQL
    End If
    
    'add new tables to table MACROTable
    sSQL = "INSERT INTO MACROTable VALUES ('Eligibility', '250', 1, 0, 0)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTable VALUES ('UniquenessCheck', '260', 1, 0, 0)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTable VALUES ('SubjectNumbering', '270', 1, 0, 0)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTable VALUES ('RSSubjectIdentifier', '', 0, 0, 0)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTable VALUES ('RSNextNumber', '', 0, 0, 0)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTable VALUES ('RSUniquenessCheck', '', 0, 0, 0)"
    MacroADODBConnection.Execute sSQL
    
    
    'add tables changes to table NewDBColumn
    sSQL = "INSERT INTO NewDBColumn VALUES (2,1,13,'TrialSubject','RegistrationStatus',null,'#NULL#')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,1,13,'MIMessage','MIMessageResponseTimeStamp',null,'#NULL#')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,1,13,'DataItemResponse','LaboratoryCode',null,'#NULL#')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,1,13,'DataItemResponseHistory','LaboratoryCode',null,'#NULL#')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,1,13,'StudyDefinition','RRServerType',1,'#NULL#')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,1,13,'StudyDefinition','RRHTTPAddress',2,'#NULL#')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,1,13,'StudyDefinition','RRUserName',3,'#NULL#')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,1,13,'StudyDefinition','RRPassword',4,'#NULL#')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,1,13,'StudyDefinition','RRProxyServer',5,'#NULL#')"
    MacroADODBConnection.Execute sSQL
    
    
    'Upgrade BuildSubVersion to [13]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '13'"
    MacroADODBConnection.Execute sSQL
    

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeData2_1from12to13", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeData2_1from17to30()
'---------------------------------------------------------------------
Dim sSQL As String
Dim oMacroDatabase As Database

    On Error GoTo ErrHandler
    
    'TA 30/1/2001 SR4131: allow zero length on TrialSubject.Gender in Access
    If goUser.Database.DatabaseType = MACRODatabaseType.Access Then
        'Using DAO
        Set oMacroDatabase = DBEngine.Workspaces(0).OpenDatabase(goUser.Database.DatabaseLocation, False, False, "MS Access;PWD=" & goUser.Database.DatabasePassword)
        oMacroDatabase.TableDefs.Refresh
        oMacroDatabase.TableDefs("TrialSubject").Fields("Gender").AllowZeroLength = True
        Set oMacroDatabase = Nothing
    End If

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeData2_1from17to30", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeData2_1from34to35()
'---------------------------------------------------------------------
Dim sSQL As String


    On Error GoTo ErrHandler

    'TA 01/03/2001: Lab data type for reporting use
    sSQL = "Insert Into DataType (DataTypeId, DataTypeName) values (6,'Laboratory Test')"
    MacroADODBConnection.Execute sSQL
    
    'update MACROControl
    Call UpGradeDataToSubVersion("35")

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeData2_1from34to35", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeSecurity2_1from39to40()
'---------------------------------------------------------------------
Dim sSQL As String


    On Error GoTo ErrHandler

    'TA 18/04/2001: new Access Create Data Views function
    sSQL = "INSERT INTO Function (FunctionCode,Function) VALUES ('F1007','Access Create Data Views')"
    SecurityADODBConnection.Execute sSQL
    sSQL = "INSERT INTO Rolefunction (RoleCode,FunctionCode) VALUES ('MacroUser','F1007')"
    SecurityADODBConnection.Execute sSQL
    
    'Upgrade BuildSubVersion to 40
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '40'"
    SecurityADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeSecurity2_1from39to40", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub


'---------------------------------------------------------------------
Private Sub UpGradeSecurity2_1from13to14()
'---------------------------------------------------------------------
Dim sSQL As String


    On Error GoTo ErrHandler
    
    ' Add new permissions
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F3024','Maintain Registration')"
    SecurityADODBConnection.Execute (sSQL)

    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F3024')"
    SecurityADODBConnection.Execute (sSQL)
    
    'Upgrade BuildSubVersion to 14
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '14'"
    SecurityADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurity2_1from13to14", "modUpgradeDatabases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub FieldNameTypeUpgrade(ByVal nDBFunction As eMACRODBFunction, _
                                ByVal nDBType As MACRODatabaseType, _
                                ByVal oConnection As ADODB.Connection, _
                                ByVal sTable As String, _
                                ByVal sVersion As String)
'---------------------------------------------------------------------
'Utility used to upgrade a table with one or more fields that have had
'a Name and/or Type change.
'
'sTable contains the table to be upgraded.
'sVersion contains the version that the table is being updated to.
'
'The upgrade happens in 4 stages:-
'   Read the contents of the specified table out into a recordset.
'   Drop the old version of the table.
'   Create a new version of the table using CreateDB.
'   Read the contents of the recordset back into the newly created version of the table.
'
'Note that if a table has had new fields added, or old fields removed, at the same time
'as other fields have had name/type changes then the adding and/or dropping should take place
'prior to this sub being called. (i.e. the old and new versions of a table need to have the
'same number of fields for this sub to work)
'
'THIS UTILITY WILL HANDLE TYPE CHANGES BETWEEN INTEGER AND LONG
'THIS UTILITY WILL HANDLE TYPE CHANGES BETWEEN AND STRINGS OF DIFFERENT LENGTHS
'THIS UTILITY WILL NOT HANDLE TYPE CHANGES THAT REQUIRE A FUNDAMENTAL TYPE CHANGE (I.E. INTEGER TO STRING)
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsFromTable As ADODB.Recordset
Dim rsToTable As ADODB.Recordset
Dim j As Long
Dim i As Integer
Dim lNumberOfRecords As Long
Dim nNumberOfFields As Integer

    On Error GoTo ErrHandler

    'Place the contents of the table into recordset rsFromTable
    sSQL = "SELECT * FROM " & sTable
    Set rsFromTable = New ADODB.Recordset
    rsFromTable.CursorLocation = adUseClient
    rsFromTable.Open sSQL, oConnection, adOpenKeyset, adLockReadOnly, adCmdText
    rsFromTable.ActiveConnection = Nothing
    
    'Drop the original version of the table
    sSQL = "DROP Table " & sTable
    oConnection.Execute sSQL
    
    'Create a new version of the table
    'Mo Morris 1/10/01, optional field to prevent CreateDB displaying messages added
    Call CreateDB(nDBFunction, nDBType, oConnection, False, False, sTable, sVersion)
    
    'Prepare recordset rsToTable to receive the contents of rsFromTable
    sSQL = "SELECT * FROM " & sTable
    Set rsToTable = New ADODB.Recordset
    rsToTable.Open sSQL, oConnection, adOpenKeyset, adLockPessimistic, adCmdText
    
    'copy the contents of rsTemp into the newly created table
    lNumberOfRecords = rsFromTable.RecordCount
    nNumberOfFields = rsFromTable.Fields.Count
    For j = 1 To lNumberOfRecords
        rsToTable.AddNew
        For i = 0 To nNumberOfFields - 1
            rsToTable.Fields(i).Value = rsFromTable.Fields(i).Value
        Next i
        rsToTable.Update
        rsFromTable.MoveNext
    Next j
    
    rsToTable.Close
    Set rsToTable = Nothing
    rsFromTable.Close
    Set rsFromTable = Nothing

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "FieldNameTypeUpgrade", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub DropDefaultConstraint(ByVal sTableName As String, _
                                    ByVal sColumnName As String)
'---------------------------------------------------------------------
'The SQL within this sub was created by Richard Weare
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sDefaultConstraintName As String

    sSQL = "SELECT name FROM sysobjects " _
        & "WHERE id =( SELECT constid " _
                    & "FROM sysconstraints " _
                    & "WHERE constid in (  SELECT id " _
                                        & "FROM sysobjects " _
                                        & "WHERE xtype = 'D' " _
                                        & "AND parent_obj = (  SELECT id " _
                                                            & "FROM sysobjects " _
                                                            & "WHERE name = '" & sTableName & "')) " _
        & "AND colid in (  SELECT colid " _
                        & "FROM syscolumns " _
                        & "WHERE name = '" & sColumnName & "' " _
                        & "AND id = (  SELECT id " _
                                    & "FROM sysobjects " _
                                    & "WHERE name = '" & sTableName & " ' " _
                                    & "AND type = 'U')) " _
                    & ")"
                    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    'If a default constraint exists the record count should be 1
    If rsTemp.RecordCount = 1 Then
        sDefaultConstraintName = rsTemp!Name
        sSQL = "ALTER Table " & sTableName & " DROP Constraint " & sDefaultConstraintName
        MacroADODBConnection.Execute sSQL
    End If

End Sub

'---------------------------------------------------------------------
Private Sub UpGradeData2_2from1to5()
'---------------------------------------------------------------------
' RJCW 26/11/01 Upgrade to database structure to include the new
' keyword table
'---------------------------------------------------------------------

Dim sSQL As String
'Dim oTableDef As TableDef
Dim oMacroDatabase As Database

On Error GoTo ErrHandler

'Create the new Keyword table
    
    
' *** Add New Tables to the Database ***
        'TA 6/12/2001: Pass through connection object rather than string so that MACROADODBConnection is up to date
        Call CreateDB(eMACRODBFunction.Data, goUser.Database.DatabaseType, MacroADODBConnection, False, False, "Keyword", "2.2.5")
        
'Insert data into new Keyword table
    
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ABSOLUTE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ACCESS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ACTION')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ADA')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ADD')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ALL')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ALLOCATE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ALTER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('AND')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ANY')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ARE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('AS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ASC')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ASSERTION')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('AT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('AUDIT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('AUTHORIZATION')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('AUTONUMBER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('AVG')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('BEGIN')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('BETWEEN')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('BIT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('BIT_LENGTH')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('BOTH')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('BY')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('BYTE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CASCADE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CASCADED')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CASE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CAST')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CATALOG')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CATEGORY')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CHAR')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CHAR_LENGTH')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CHARACTER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CHARACTER_LENGTH')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CHECK')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CLINICALTRIALID')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CLOSE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CLUSTER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('COALESCE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('COLLATE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('COLLATION')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('COLUMN')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('COMMENT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('COMMIT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('COMPRESS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CONNECT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CONNECTION')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CONSTRAINT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CONSTRAINTS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CONTINUE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CONVERT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CORRESPONDING')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('COUNT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CREATE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CRFPAGECYCLENUMBER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CRFPAGEID')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CROSS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CURRENCY')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CURRENT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CURRENT_DATE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CURRENT_TIME')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CURRENT_TIMESTAMP')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CURRENT_USER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CURSOR')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('CYCLE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DATE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DAY')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DEALLOCATE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DEC')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DECIMAL')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DECLARE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DEFAULT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DEFERRABLE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DEFERRED')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DELETE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DESC')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DESCRIBE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DESCRIPTOR')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DIAGNOSTICS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DISCONNECT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DISTINCT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DOMAIN')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DOUBLE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DROP')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DYNAMIC')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ELSE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('END')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('END-EXEC')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ESCAPE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('EXCEPT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('EXCEPTION')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('EXCLUSIVE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('EXEC')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('EXECUTE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('EXISTS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('EXTERNAL')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('EXTRACT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FALSE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FETCH')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FILE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FIRST')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FLOAT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FOR')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FOREIGN')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FORM')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FORTRAN')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FOUND')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FROM')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FULL')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('GET')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('GLOBAL')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('GO')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('GOTO')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('GRANT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('GROUP')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('HAVING')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('HOUR')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('HYPERLINK')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('IDENTIFIED')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('IDENTITY')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('IMMEDIATE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('IN')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INCLUDE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INCREMENT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INDEX')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INDICATOR')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INITIAL')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INITIALIZATION')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INITIALLY')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INNER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INPUT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INSENSITIVE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INSERT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INTEGER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INTERSECT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INTERVAL')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INTO')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('IS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ISOLATION')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('JOIN')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('KEY')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LANGUAGE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LAST')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LEADING')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LEFT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LEVEL')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LIKE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LOCAL')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LOCK')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LONG')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LOWER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MATCH')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MAX')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MAXEXTENTS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MEMO')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('META_PREDICATE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MIN')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MINUS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MINUTE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MOD')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MODE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MODIFY')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MODULE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MONTH')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MULTIFILE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NAME')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NAMES')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NATIONAL')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NATURAL')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NCHAR')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NEXT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NO')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NOAUDIT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NOCOMPRESS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NONE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NOSPY')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NOT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NOWAIT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NULL')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NULLIF')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NUMBER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NUMERIC')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('OCTET_LENGTH')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('OF')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('OFFLINE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ON')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ONE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ONLINE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ONLY')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('OPEN')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('OPTION')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('OR')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ORDER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('OUTER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('OUTPUT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('OVERLAPS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PAD')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PARTIAL')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PASCAL')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PASSWORD')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PCTFREE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PERSONID')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('POSITION')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PRECISION')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PREPARE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PRESERVE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PRIMARY')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PRIOR')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PRIVILEGES')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PROCEDURE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('PUBLIC')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('QUESTION')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('RANDOMISATION')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('RAW')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('READ')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('REAL')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('REFERENCES')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('REGISTRATION')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('RELATIVE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('RENAME')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('RESOURCE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('RESTRICT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('REVOKE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('RIGHT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ROLLBACK')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ROW')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ROWID')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ROWLABEL')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ROWNUM')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ROWS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SCHEMA')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SCROLL')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SECOND')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SECTION')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SELECT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SESSION')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SESSION_USER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SET')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SHARE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SITE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SIZE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SMALLINT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SOME')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SPACE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SPY')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SQL')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SQLCA')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SQLCODE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SQLERROR')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SQLSTATE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SQLWARNING')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('START')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('STATUS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SUBSTRING')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SUCCESSFUL')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SUM')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SYNONYM')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SYSDATE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SYSTEM_USER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TABLE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TEMP')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TEMPORARY')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TEXT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('THEN')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TIME')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TIMESTAMP')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TIMEZONE_HOUR')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TIMEZONE_MINUTE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TINYINT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TO')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TRAILING')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TRANSACTION')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TRANSLATE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TRANSLATION')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TRIGGER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TRIM')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TRUE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('UID')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('UNION')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('UNIQUE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('UNKNOWN')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('UPDATE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('UPPER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('USAGE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('USER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('USING')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VALIDATE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VALUE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VALUES')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VARCHAR')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VARCHAR2')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VARYING')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VIEW')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VISIT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VISITCYCLENUMBER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VISITID')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('VOLATILE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('WHEN')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('WHENEVER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('WHERE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('WITH')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('WORK')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('WRITE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('YEAR')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('YES')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('YESNO')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ZONE')"
        MacroADODBConnection.Execute sSQL
        'REM 10/05/02 - added new keywords
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ABS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('AFTER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('AVERAGE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('BEFORE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DATE_AND_TIME')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DATE_DIFF')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DATEDIFF')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DATENOW')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DATEOF')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DATEPART')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DATETIME')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DAYS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('DURING')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('EVERY')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('FORMAT_DATE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('HOURS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('IF')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('INCLUDES')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('IS_KNOWN')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ISAFTER')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ISBEFORE')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ISDURING')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ISKNOWN')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('LEN')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MINUTES')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('MONTHS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NETSUPPORT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NOT_ONEOF')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('NOW')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('ONEOF')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('RESULT_OF')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SECONDS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('SQRT')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TIME_DIFF')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('TIMENOW')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('WEEK')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('WEEKS')"
        MacroADODBConnection.Execute sSQL
        sSQL = "INSERT INTO Keyword (Keyword) VALUES ('YEARS')"
        MacroADODBConnection.Execute sSQL

        
        'insert table information into MACROTABLE for Keyword
        sSQL = "INSERT INTO MACROTable (TableName, SegmentId, STYDEF, PATRSP, LABDEF) "
        sSQL = sSQL & " VALUES ('Keyword', '', 0, 0, 0)"
        MacroADODBConnection.Execute sSQL
    
    Exit Sub

ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeData2_2from1to5", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'---------------------------------------------------------------------
Public Sub UpGradeData2_2from5to10()
'---------------------------------------------------------------------
' REM 15/04/02 Upgrade to MACRO 2.2 database structure to include the new
' ArezzoToken and AutoImport tables and a new column called HadValue to the
' DataItemResponse and DataItemResponseHistory tables and sql to insert the correct
' HadValue into the column
'---------------------------------------------------------------------
Dim sSQL As String
    
    On Error GoTo ErrHandler

    ' *** Add New Tables to the Database ***
    Call CreateDB(eMACRODBFunction.Data, goUser.Database.DatabaseType, MacroADODBConnection, False, False, "ArezzoToken", "2.2.10")

    Call CreateDB(eMACRODBFunction.Data, goUser.Database.DatabaseType, MacroADODBConnection, False, False, "AutoImportControl", "2.2.10")

    ' *** Add New Tables to the table MACROTable ***
    sSQL = "INSERT INTO MACROTable VALUES ('ArezzoToken','',0,0,0)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTable VALUES ('AutoImportControl','',0,0,0)"
    MacroADODBConnection.Execute sSQL

    ' *** Add New Columns ***
    'Added 1 column to DataItemResponse: HadValue
    'Added 1 column to DataItemResponseHistory: HadValue
    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.Access
        sSQL = "ALTER Table DataItemResponse ADD COLUMN HadValue SMALLINT"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table DataItemResponseHistory ADD COLUMN HadValue SMALLINT"
        MacroADODBConnection.Execute sSQL
    
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
        sSQL = "ALTER Table DataItemResponse ADD HadValue SMALLINT"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table DataItemResponseHistory ADD HadValue SMALLINT"
        MacroADODBConnection.Execute sSQL
        
    Case MACRODatabaseType.Oracle80
        sSQL = "ALTER Table DataItemResponse ADD HadValue NUMBER(6)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table DataItemResponseHistory ADD HadValue NUMBER(6)"
        MacroADODBConnection.Execute sSQL
        
    End Select

    ' *** Drop Columns ***
    'First drop the default constraints on the columns for SQL server
    If goUser.Database.DatabaseType = MACRODatabaseType.sqlserver Or goUser.Database.DatabaseType = MACRODatabaseType.SQLServer70 Then
        Call DropDefaultConstraint("ClinicalTrial", "ActualRecruitment")
        Call DropDefaultConstraint("Trialsite", "TrialSiteActualRecruitment")
    End If
    
    'Drop ActualRecruitment column from ClinicalTrail table
    'Drop TrialSiteActualRecruitment column from TrialSite table
    sSQL = "ALTER Table ClinicalTrial DROP Column ActualRecruitment"
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table TrialSite DROP Column TrialSiteActualRecruitment"
    MacroADODBConnection.Execute sSQL

    ' *** Add values to the new HadValue columns ***
    'Update DataItemResponseHistory table
    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.Access
    sSQL = "Update DataItemResponseHistory a" _
        & " Set HadValue = 0" _
        & " WHERE   (a.ResponseValue IS NULL AND NOT EXISTS (SELECT  * FROM DataItemResponseHistory b" _
        & "                                                  WHERE b.ResponseValue IS NOT Null" _
        & "                                                  AND b.ClinicalTrialId = a.ClinicalTrialId" _
        & "                                                  AND b.TrialSite = a.TrialSite" _
        & "                                                  AND b.PersonId = a.PersonId" _
        & "                                                  AND b.ResponseTaskId = a.ResponseTaskId" _
        & "                                                  AND b.ResponseTimeStamp < a.ResponseTimeStamp))"
        MacroADODBConnection.Execute sSQL
    
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
    sSQL = "Update DataItemResponseHistory" _
        & " Set HadValue = 0" _
        & " FROM    DataItemResponseHistory a" _
        & " WHERE   (a.ResponseValue IS NULL AND NOT EXISTS (SELECT  * FROM DataItemResponseHistory b" _
        & "                                                  WHERE b.ResponseValue IS NOT Null" _
        & "                                                  AND b.ClinicalTrialId = a.ClinicalTrialId" _
        & "                                                  AND b.TrialSite = a.TrialSite" _
        & "                                                  AND b.PersonId = a.PersonId" _
        & "                                                  AND b.ResponseTaskId = a.ResponseTaskId" _
        & "                                                  AND b.ResponseTimeStamp < a.ResponseTimeStamp))"
        MacroADODBConnection.Execute sSQL
        
    Case MACRODatabaseType.Oracle80
    sSQL = "Update DataItemResponseHistory a" _
        & " Set HadValue = 0" _
        & " WHERE   (a.ResponseValue IS NULL AND NOT EXISTS (SELECT  * FROM DataItemResponseHistory b" _
        & "                                                  WHERE b.ResponseValue IS NOT Null" _
        & "                                                  AND b.ClinicalTrialId = a.ClinicalTrialId" _
        & "                                                  AND b.TrialSite = a.TrialSite" _
        & "                                                  AND b.PersonId = a.PersonId" _
        & "                                                  AND b.ResponseTaskId = a.ResponseTaskId" _
        & "                                                  AND b.ResponseTimeStamp < a.ResponseTimeStamp))"
        MacroADODBConnection.Execute sSQL
        
    End Select

    sSQL = "Update DataItemResponseHistory" _
        & " SET HadValue = 1 WHERE HadValue IS NULL"
        MacroADODBConnection.Execute sSQL


    'Update DataItemResponse table by copying value from DataItemResponseHistory table
    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.Access
    sSQL = "Update DataItemResponse, DataItemResponseHistory" _
        & " Set DataItemResponse.HadValue = DataItemResponseHistory.HadValue" _
        & " WHERE DataItemResponse.ClinicalTrialId = DataItemResponseHistory.ClinicalTrialId" _
        & " AND DataItemResponse.TrialSite = DataItemResponseHistory.TrialSite" _
        & " AND DataItemResponse.PersonId = DataItemResponseHistory.PersonId" _
        & " AND DataItemResponse.ResponseTaskId = DataItemResponseHistory.ResponseTaskId" _
        & " AND DataItemResponse.ResponseTimeStamp = DataItemResponseHistory.ResponseTimeStamp"
        MacroADODBConnection.Execute sSQL

    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
    sSQL = "Update DataItemResponse" _
        & " Set DataItemResponse.HadValue = DataItemResponseHistory.HadValue" _
        & " FROM DataItemResponse, DataItemResponseHistory" _
        & " WHERE DataItemResponse.ClinicalTrialId = DataItemResponseHistory.ClinicalTrialId" _
        & " AND DataItemResponse.TrialSite = DataItemResponseHistory.TrialSite" _
        & " AND DataItemResponse.PersonId = DataItemResponseHistory.PersonId" _
        & " AND DataItemResponse.ResponseTaskId = DataItemResponseHistory.ResponseTaskId" _
        & " AND DataItemResponse.ResponseTimeStamp = DataItemResponseHistory.ResponseTimeStamp"
        MacroADODBConnection.Execute sSQL

    Case MACRODatabaseType.Oracle80
     sSQL = "UPDATE DataItemResponse SET DataItemResponse.Hadvalue = (Select DataItemResponseHistory.HadValue" _
        & " From DataItemResponseHistory" _
        & " Where DataItemResponse.ClinicalTrialId = DataItemResponseHistory.ClinicalTrialId" _
        & " AND DataItemResponse.TrialSite = DataItemResponseHistory.TrialSite" _
        & " AND DataItemResponse.PersonId = DataItemResponseHistory.PersonId" _
        & " AND DataItemResponse.ResponseTaskId = DataItemResponseHistory.ResponseTaskId" _
        & " AND DataItemResponse.ResponseTimeStamp = DataItemResponseHistory.ResponseTimeStamp)"
        MacroADODBConnection.Execute sSQL
    
    End Select
    
    ' *** Add the new column names to the NewDBColumn table ***
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,10,'DataItemResponse','HadValue',null,'1','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,10,'DataItemResponseHistory','HadValue',null,'1','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    
    ' *** Add the dropped columns to the NewDBColumn table ***
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,10,'ClinicalTrial','ActualRecruitment',null,'#NULL#','DROPCOLUMN',8)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (2,2,10,'TrialSite','TrialSiteActualRecruitment',null,'#NULL#','DROPCOLUMN',3)"
    MacroADODBConnection.Execute sSQL
    
    '*** Add new Index ***
    sSQL = "CREATE INDEX IDX_VI_SECONDARYKEY "
    sSQL = sSQL & "ON VisitInstance ( ClinicalTrialId, TrialSite, PersonId, VisitId, VisitCycleNumber )"
    MacroADODBConnection.Execute sSQL
    
    'Update BuildSubVersion to [10]
    sSQL = "UPDATE MACROControl Set BuildSubVersion = '10'"
    MacroADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeData2_2from5to10", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'---------------------------------------------------------------------
Private Sub UpGradeData2_2from14to15()
'---------------------------------------------------------------------
' REM 10/06/02
' Insert Default values into the AutoImportControl table
'---------------------------------------------------------------------
Dim sSQL As String

    'if there is already a row in the table then
    On Error Resume Next
    sSQL = "INSERT INTO AutoImportControl VALUES (1,'START',60)"
    MacroADODBConnection.Execute sSQL

End Sub

'---------------------------------------------------------------------
Private Sub UpgradeSecurity2_2From15to16()
'---------------------------------------------------------------------
'Mo Morris 11/6/2002, Add new Function code 'F1008' to table Function
'   and table Rolefunction for RoleCode 'MacroUser'
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "INSERT INTO Function (FunctionCode,Function) VALUES ('F1008','Access Query Module')"
    SecurityADODBConnection.Execute sSQL
    sSQL = "INSERT INTO Rolefunction (RoleCode,FunctionCode) VALUES ('MacroUser','F1008')"
    SecurityADODBConnection.Execute sSQL
    
'TA 13/06/02 CBB 2.2.13.43 F5006 (view reports) function reinserted into Function and RoleFunction tables
    sSQL = "Insert into Function values ('F5006','View reports')"
    SecurityADODBConnection.Execute sSQL
    
    sSQL = "insert into RoleFunction values ('MacroUser','F5006')"
    SecurityADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpgradeSecurity2_2From15to16", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'---------------------------------------------------------------------
Private Sub UpgradeData2_2from18to19()
'---------------------------------------------------------------------
' MLM 28/06/02: Added. Add indexes to TrialSubject to improve performance.
'---------------------------------------------------------------------

Dim sSQL As String

    sSQL = "CREATE INDEX IDX_TS_TRIALSITE "
    sSQL = sSQL & "ON TrialSubject ( TrialSite )"
    MacroADODBConnection.Execute sSQL

    sSQL = "CREATE INDEX IDX_TS_PERSONID "
    sSQL = sSQL & "ON TrialSubject ( PersonId )"
    MacroADODBConnection.Execute sSQL


End Sub

'---------------------------------------------------------------------
Private Sub UpgradeDataDatabase3_0()
'---------------------------------------------------------------------
' Upgrade a 3.0 MACRO db
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim sSQL As String
Dim sMacroVersion As String
Dim sBuildSubVersion As String
Dim sMsg As String
Dim sScriptPrefix As String
Dim vSQL As String
Dim sUpgradePath As String
Dim lBuildSubVersion As Long

    On Error GoTo ErrHandler
    
'    'check if MSDE Upgrade is needed
'    If Not MSDEUpgrade Then
'        Call DialogError("An Access database cannot be used in MACRO 3.0.  MACRO will now close down")
'        'code to close down MACRO
'        Call ExitMACRO
'        Call MACROEnd
'    End If
'
    'will not get here if we have an access database
    sSQL = "SELECT * FROM MACROControl"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        MsgBox ("Your Macro database is not valid. Macro is being closed down.")
        ExitMACRO
        MACROEnd
    End If
    
    sBuildSubVersion = rsTemp![BuildSubVersion]
    sMacroVersion = rsTemp![MACROVersion]
    rsTemp.Close
    Set rsTemp = Nothing

    ' Check for version 2.2
    If sMacroVersion = "2.2" Then
        sMsg = "You are about to upgrade your MACRO database from 2.2 to 3.0. Do you wish to continue?"
        Select Case MsgBox(sMsg, vbQuestion + vbYesNo, gsDIALOG_TITLE)
            Case vbYes
                Call UpGradeData2_2to3_0_7
                sMacroVersion = "3.0"
                sBuildSubVersion = "7"
            'Upgrade MacroVersion to [3.0]
                sSQL = "UPDATE MACROControl Set MacroVersion = '3.0'"
                MacroADODBConnection.Execute sSQL
                'update subversion in database
                Call UpGradeDataToSubVersion(sBuildSubVersion)
            Case vbNo
                Call ExitMACRO
                Call MACROEnd
        End Select
    End If
    
    If sMacroVersion <> "3.0" Then
        MsgBox ("Your MACRO database is not valid. MACRO is being closed down.")
        ExitMACRO
        MACROEnd
    End If
    
    'TODO remove this when we release
    ' anything below 7 needs manual upgrade
    If Val(sBuildSubVersion) < 7 Then
        DialogInformation "A manual upgrade is required for a " & sMacroVersion & "." & sBuildSubVersion & " database"
        ExitMACRO
        MACROEnd
    End If
    
    'ADD NEW 3.0 UPGRADES HERE
    'REM 23/08/02 - Upgrade from 7 to 8
    If sBuildSubVersion = "7" Then
        sBuildSubVersion = "8"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
        Call UpgradeData3_0from7to8
    End If

    'REM 29/08/02 - Upgrade from 8 to 9
    If sBuildSubVersion = "8" Then
        sBuildSubVersion = "9"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
        Call UpgradeData3_0from8to9
    End If
    
    'TA 16/009/2002 - Upgrade from 9 to 12
    If sBuildSubVersion = "9" Then
        sBuildSubVersion = "12"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
        Call UpgradeData3_0from9to12
    End If

    'TA 16/009/2002 - Upgrade from 9 to 12
    If sBuildSubVersion = "12" Then
        sBuildSubVersion = "13"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
        Call UpgradeData3_0from12to13
    End If
    
    'REM 08/10/02 - Upgrade from 13 to 14
    If sBuildSubVersion = "13" Then
        sBuildSubVersion = "14"
        Call UpGradeDataToSubVersion(sBuildSubVersion)
        Call Upgrade3_0from13to14
    End If
    
    'REM 08/10/02 - Upgrade from 17 to 18
    If (Val(sBuildSubVersion) >= 14) And (Val(sBuildSubVersion) < 18) Then
        sBuildSubVersion = "18"
        Call UpGradeData3_0from14to18
        Call UpGradeDataToSubVersion(sBuildSubVersion)
    End If
    
    sUpgradePath = App.Path & "\Database Scripts\Upgrade Database\"
    
    'set prefix of cript file according to dbtype
    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.Oracle80: sScriptPrefix = "ORA"
    Case Else: sScriptPrefix = "MSSQL"
    End Select
    
    For lBuildSubVersion = 18 To (CURRENT_SUBVERSION - 1) 'always 1 less than build number
        If (Val(sBuildSubVersion) = lBuildSubVersion) Then
            sBuildSubVersion = CStr(lBuildSubVersion + 1)
            ExecuteMultiLineSQL MacroADODBConnection, _
                                StringFromFile(sUpgradePath & sScriptPrefix & "_30_" & CStr(lBuildSubVersion) & "To" & sBuildSubVersion & ".sql")
        End If
    Next

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpgradeDataDatabase3_0", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select


End Sub

'---------------------------------------------------------------------
Private Sub UpgradeData3_0from7to8()
'---------------------------------------------------------------------
'REM 23/08/07
'Upgrade 3.0.7 to 3.0.8
'Create 2 new tables: StudyVersion, MACRODBSetting
'Alter 3 tables: Message tabel add new column MessageRecievedTimeStamp
'                TrialSite table add column Version
'                Site table add column SiteLocation
'TA 24/08/2002: Added CRFElement.elementuse and StudyVisitCRFPAge.eFormUse
'MLM 28/08/02: Changed from eFormUSe to EFormUse
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsCount As ADODB.Recordset
Dim nCount As Integer
Dim sIntegerDataType As String
Dim sAlterColumn As String

On Error GoTo ErrLabel

    ' *** Add New Tables to the Database ***
    Call CreateDB(eMACRODBFunction.Data, goUser.Database.DatabaseType, MacroADODBConnection, False, False, "StudyVersion", "3.0.8")
    Call CreateDB(eMACRODBFunction.Data, goUser.Database.DatabaseType, MacroADODBConnection, False, False, "MACRODBSetting", "3.0.8")
    
    ' *** Add New Tables to the table MACROTable ***
    sSQL = "INSERT INTO MACROTable VALUES ('StudyVersion','',0,0,0)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTable VALUES ('MACRODBSetting','',0,0,0)"
    MacroADODBConnection.Execute sSQL

    ' *** Add New Columns ***
    'Added 1 column to Message table: MessageReceivedTimeStamp
    'Added 1 column to Site table: SiteLocation
    'Added 1 column to TrialSite table: StudyVersion
    Select Case goUser.Database.DatabaseType
'    Case MACRODatabaseType.Access
'        sSQL = "ALTER Table Message ADD COLUMN MessageReceivedTimeStamp DOUBLE"
'        MacroADODBConnection.Execute sSQL
'        sSQL = "ALTER Table Site ADD COLUMN SiteLocation SMALLINT"
'        MacroADODBConnection.Execute sSQL
'        sSQL = "ALTER Table TrialSite ADD COLUMN StudyVersion SMALLINT"
'        MacroADODBConnection.Execute sSQL
    
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
        sSQL = "ALTER Table Message ADD MessageReceivedTimeStamp NUMERIC(16,10)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table Site ADD SiteLocation SMALLINT"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table TrialSite ADD StudyVersion SMALLINT"
        MacroADODBConnection.Execute sSQL
        
    Case MACRODatabaseType.Oracle80
        sSQL = "ALTER Table Message ADD MessageReceivedTimeStamp NUMBER(16,10)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table Site ADD SiteLocation NUMBER(6)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table TrialSite ADD StudyVersion NUMBER(6)"
        MacroADODBConnection.Execute sSQL
        
    End Select

    ' *** Add the new column names to the NewDBColumn table ***
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,8,'Message','MessageReceivedTimeStamp',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,8,'Site','SiteLocation',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,8,'TrialSite','StudyVersion',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    
    '*** Add defalts for existing data into MACRODBSetting table ***
    Set rsCount = New ADODB.Recordset
    
    sSQL = "SELECT COUNT(*) FROM TrialOffice"
    rsCount.Open sSQL, MacroADODBConnection

    nCount = rsCount.Fields(0).Value
    
    rsCount.Close
    Set rsCount = Nothing
    
    If nCount > 0 Then
        sSQL = "INSERT INTO MACRODBSetting(SettingSection, SettingKey,SettingValue) VALUES ('datatransfer','dbtype','site')"
    Else
        sSQL = "INSERT INTO MACRODBSetting(SettingSection, SettingKey,SettingValue) VALUES ('datatransfer','dbtype','server')"
    End If
    
    MacroADODBConnection.Execute sSQL
    
    
    'TA 24/08/2002: EForm and Visit Date Changes
    'set up db specific keywqords
    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.Oracle80
        sIntegerDataType = "NUMBER (6)"
        sAlterColumn = "MODIFY"
    Case Else
        sIntegerDataType = "SMALLINT"
        sAlterColumn = "ALTER COLUMN"
    End Select
    
    ' *** Add New Columns ***
    'Added 1 column to CRFElement table: ElementUse - default 0
    'Added 1 column to StudyVistCRFPage table: eFormUse - default 0
    sSQL = "ALTER Table CRFElement ADD ElementUse " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table StudyVisitCRFPage ADD EFormUse " & sIntegerDataType
    MacroADODBConnection.Execute sSQL

    sSQL = "UPDATE CRFElement SET ElementUse = 0 WHERE ElementUse IS NULL"
    MacroADODBConnection.Execute sSQL
    sSQL = "UPDATE StudyVisitCRFPage SET eFormUse = 0 WHERE eFormUse IS NULL"
    MacroADODBConnection.Execute sSQL

    'put not nulls on appropriate columns
    sSQL = "ALTER TABLE CRFElement " & sAlterColumn & " ElementUse " & sIntegerDataType & " NOT NULL"
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER TABLE StudyVisitCRFPage " & sAlterColumn & " eFormUse " & sIntegerDataType & " NOT NULL"
    MacroADODBConnection.Execute sSQL


    ' *** Add the new column names to the NewDBColumn table ***
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,8,'CRFElement','ElementUse',null,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,8,'StudyVisitCRFPage','eFormUse',null,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL

        
Exit Sub
    
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modUpgradeDatabases.UpgradeData3_0from7to8"
End Sub

'---------------------------------------------------------------------
Private Sub UpgradeData3_0from8to9()
'---------------------------------------------------------------------
'REM 29/08/02
'Input values into the new DiscrepancyStatus and SDVStatus columns in the
' TrialSubject, VisitInstance, CRFPageInstance, DataItemResponse tables
'---------------------------------------------------------------------
Dim sSQL As String
Dim bSQLServer As Boolean
Dim vOld As Variant
Dim vNew As Variant
Dim sElse As String
Dim i As Integer
Dim sStatusColumn As String
Dim nMIMessageType As Integer
Dim sAlterColumn As String
Dim sIntegerDataType As String

    On Error GoTo ErrLabel

    'Check database type
    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.Oracle80
        sIntegerDataType = "NUMBER (6)"
        sAlterColumn = "MODIFY"
        bSQLServer = False
    Case Else
        sIntegerDataType = "SMALLINT"
        sAlterColumn = "ALTER COLUMN"
        bSQLServer = True
    End Select

    'Loop through the following code twice to update both the DiscrepancyStaus and SDVStatus columns
    For i = 1 To 2
        
        If i = 1 Then 'update the DiscrepancyStaus column
            sStatusColumn = "DiscrepancyStatus"
            nMIMessageType = 0
            'Discrepancys - raised, responded, closed
            vOld = Array("0", "1", "2")
            vNew = Array("30", "20", "10")
        ElseIf i = 2 Then 'update the SDVStaus column
            sStatusColumn = "SDVStatus"
            nMIMessageType = 3
            'SDV's - planned , complete(done)
            vOld = Array("0", "2")
            vNew = Array("30", "20")
        End If
        
        'set the value to 0 in all other cases
        sElse = "0"
        
        'update the TrialSubject table
        sSQL = "Update TrialSubject set " & sStatusColumn & " = (" _
             & " SELECT MAX(" & CreateDecodeSQL(bSQLServer, "MIMessageStatus", vOld, vNew, sElse) & ")" _
             & " From MIMESSAGE, ClinicalTrial" _
             & " Where MIMessageType = " & nMIMessageType _
             & " AND MIMessageHistory = 0" _
             & " AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName" _
             & " AND MIMessageSite = TrialSubject.TrialSite" _
             & " AND MIMessagePersonId = TrialSubject.PersonId" _
             & " AND ClinicalTrial.ClinicalTrialId = TrialSubject.ClinicalTrialId )"
        MacroADODBConnection.Execute sSQL
        
        'Update all the Nulls to 0
        sSQL = "Update TrialSubject SET " & sStatusColumn & " = 0 Where " & sStatusColumn & " IS NULL"
        MacroADODBConnection.Execute sSQL
        
        'update the VisitInstance table
        sSQL = "Update VisitInstance set " & sStatusColumn & " = (" _
             & " SELECT MAX(" & CreateDecodeSQL(bSQLServer, "MIMessageStatus", vOld, vNew, sElse) & ")" _
             & " From MIMESSAGE, ClinicalTrial" _
            & " Where MIMessageType = " & nMIMessageType _
             & " AND MIMessageHistory = 0" _
             & " AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName" _
             & " AND MIMessageSite = VisitInstance.TrialSite" _
             & " AND MIMessagePersonId = VisitInstance.PersonId" _
             & " AND MIMessageVisitId = VisitInstance.VisitId" _
             & " AND MIMessageVisitCycle = VisitInstance.VisitCycleNumber" _
             & " AND ClinicalTrial.ClinicalTrialId = VisitInstance.ClinicalTrialId )"
        MacroADODBConnection.Execute sSQL
        
        'Update all the Nulls to 0
        sSQL = "Update VisitInstance SET " & sStatusColumn & " = 0 Where " & sStatusColumn & " IS NULL"
        MacroADODBConnection.Execute sSQL
    
        'update the CRFPageInstance table
        sSQL = "Update CRFPageInstance set " & sStatusColumn & " = (" _
             & " SELECT MAX(" & CreateDecodeSQL(bSQLServer, "MIMessageStatus", vOld, vNew, sElse) & ")" _
             & " From MIMESSAGE, ClinicalTrial" _
            & " Where MIMessageType = " & nMIMessageType _
             & " AND MIMessageHistory = 0" _
             & " AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName" _
             & " AND MIMessageSite = CRFPageInstance.TrialSite" _
             & " AND MIMessagePersonId = CRFPageInstance.PersonId" _
             & " AND MIMessageCRFPageTaskId = CRFPageInstance.CRFPageTaskId" _
             & " AND ClinicalTrial.ClinicalTrialId = CRFPageInstance.ClinicalTrialId )"
        MacroADODBConnection.Execute sSQL
    
        'Update all the Nulls to 0
        sSQL = "Update CRFPageInstance SET " & sStatusColumn & " = 0 Where " & sStatusColumn & " IS NULL"
        MacroADODBConnection.Execute sSQL
    
        'update the DataItemResponse table
        sSQL = "Update DataItemResponse set " & sStatusColumn & " = (" _
             & " SELECT MAX(" & CreateDecodeSQL(bSQLServer, "MIMessageStatus", vOld, vNew, sElse) & ")" _
             & " From MIMESSAGE, ClinicalTrial" _
            & " Where MIMessageType = " & nMIMessageType _
             & " AND MIMessageHistory = 0" _
             & " AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName" _
             & " AND MIMessageSite = DataItemResponse.TrialSite" _
             & " AND MIMessagePersonId = DataItemResponse.PersonId" _
             & " AND MIMessageResponseTaskId = DataItemResponse.ResponseTaskId" _
             & " AND MIMessageResponseCycle = DataItemResponse.RepeatNumber" _
             & " AND ClinicalTrial.ClinicalTrialId = DataItemResponse.ClinicalTrialId )"
        MacroADODBConnection.Execute sSQL
    
        'Update all the Nulls to 0
        sSQL = "Update DataItemResponse SET " & sStatusColumn & " = 0 Where " & sStatusColumn & " IS NULL"
        MacroADODBConnection.Execute sSQL
        
    Next
    
    'update the DataItemResponse table
    sSQL = "Update DataItemResponse set NoteStatus = (" _
         & " SELECT count(*)" _
         & " From MIMESSAGE, ClinicalTrial" _
        & " Where MIMessageType = 2" _
         & " AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName" _
         & " AND MIMessageSite = DataItemResponse.TrialSite" _
         & " AND MIMessagePersonId = DataItemResponse.PersonId" _
         & " AND MIMessageResponseTaskId = DataItemResponse.ResponseTaskId" _
         & " AND MIMessageResponseCycle = DataItemResponse.RepeatNumber" _
         & " AND ClinicalTrial.ClinicalTrialId = DataItemResponse.ClinicalTrialId )"
    MacroADODBConnection.Execute sSQL

    'Update all the non zeros to 1
    sSQL = "Update DataItemResponse SET NoteStatus = 1 Where NoteStatus > 0"
    MacroADODBConnection.Execute sSQL
    
    
    'put not nulls on appropriate columns
    sSQL = "ALTER TABLE DataItemResponse " & sAlterColumn & " ChangeCount " & sIntegerDataType & " NOT NULL"
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER TABLE DataItemResponse " & sAlterColumn & " DiscrepancyStatus " & sIntegerDataType & " NOT NULL"
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER TABLE DataItemResponse " & sAlterColumn & " SDVStatus " & sIntegerDataType & " NOT NULL"
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER TABLE DataItemResponse " & sAlterColumn & " NoteStatus " & sIntegerDataType & " NOT NULL"
    MacroADODBConnection.Execute sSQL
    
    'do not allow nulls
    sSQL = "ALTER TABLE CRFPageInstance " & sAlterColumn & " DiscrepancyStatus " & sIntegerDataType & " NOT NULL"
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER TABLE CRFPageInstance " & sAlterColumn & " SDVStatus " & sIntegerDataType & " NOT NULL"
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER TABLE CRFPageInstance " & sAlterColumn & " NoteStatus " & sIntegerDataType & " NOT NULL"
    MacroADODBConnection.Execute sSQL
    
    'do not allow nulls
    sSQL = "ALTER TABLE VisitInstance " & sAlterColumn & " DiscrepancyStatus " & sIntegerDataType & " NOT NULL"
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER TABLE VisitInstance " & sAlterColumn & " SDVStatus " & sIntegerDataType & " NOT NULL"
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER TABLE VisitInstance " & sAlterColumn & " NoteStatus " & sIntegerDataType & " NOT NULL"
    MacroADODBConnection.Execute sSQL
    
    sSQL = "ALTER TABLE TrialSubject " & sAlterColumn & " DiscrepancyStatus " & sIntegerDataType & " NOT NULL"
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER TABLE TrialSubject " & sAlterColumn & " SDVStatus " & sIntegerDataType & " NOT NULL"
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER TABLE TrialSubject " & sAlterColumn & " NoteStatus " & sIntegerDataType & " NOT NULL"
    MacroADODBConnection.Execute sSQL
    
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modUpgradeDatabases.UpgradeData3_0from8to9"
End Sub

'---------------------------------------------------------------------
Private Sub UpgradeData3_0from9to12()
'---------------------------------------------------------------------
' RS 16/9/2002: Include new Timezone columns
'---------------------------------------------------------------------
Dim bSQLServer As Boolean
Dim sAlterColumn As String
Dim sIntegerDataType As String
Dim sTimestampDataType As String
Dim sSQL As String

    On Error GoTo ErrLabel
    
    'Check database type
    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.Oracle80
        sIntegerDataType = "NUMBER (6)"
    Case Else
        sIntegerDataType = "SMALLINT"
    End Select

    'alter tables
    sSQL = "ALTER Table CRFElement ADD DisplayLength " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table StudyVisit ADD Repeating " & sIntegerDataType
    MacroADODBConnection.Execute sSQL

    'add to new dbcolumn for imp/exp between versions
    sSQL = "INSERT into NewDbColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) values (3,0,10,'CRFElement','DisplayLength',1,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT into NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) values (3,0,10,'StudyVisit','Repeating',1,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL


Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modUpgradeDatabases.UpgradeData3_0from9to12"
    
End Sub

'---------------------------------------------------------------------
Private Sub UpgradeData3_0from12to13()
'---------------------------------------------------------------------
' RS 16/9/2002: Include new Timezone columns
'---------------------------------------------------------------------
Dim bSQLServer As Boolean
Dim sAlterColumn As String
Dim sIntegerDataType As String
Dim sTimestampDataType As String
Dim sSQL As String

    On Error GoTo ErrLabel
    
    'Check database type
    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.Oracle80
        sIntegerDataType = "NUMBER (6)"
        sAlterColumn = "MODIFY"
        sTimestampDataType = "DECIMAL(16,10)"
        bSQLServer = False
    Case Else
        sIntegerDataType = "SMALLINT"
        sAlterColumn = "ALTER COLUMN"
        sTimestampDataType = "DECIMAL(16,10)"
        bSQLServer = True
    End Select
    
    ' Add all TimeZone Columns
    sSQL = "ALTER Table TrialSubject ADD ImportTimestamp_TZ " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table VisitInstance ADD ImportTimestamp_TZ " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table CRFPageInstance ADD ImportTimestamp_TZ " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table QGroupInstance ADD ImportTimestamp_TZ " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table DataItemResponse ADD ResponseTimestamp_TZ " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table DataItemResponse ADD ImportTimestamp_TZ " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table DataItemResponse ADD DatabaseTimestamp " & sTimestampDataType
    MacroADODBConnection.Execute sSQL
    'set default value of 0
    sSQL = "Update DataItemResponse Set DatabaseTimeStamp = 0 where DatabaseTimeStamp is null"
    MacroADODBConnection.Execute sSQL
    
    sSQL = "ALTER Table DataItemResponse ADD DatabaseTimestamp_TZ " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table DataItemResponseHistory ADD ResponseTimestamp_TZ " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table DataItemResponseHistory ADD ImportTimestamp_TZ " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table DataItemResponseHistory ADD DatabaseTimestamp " & sTimestampDataType
    MacroADODBConnection.Execute sSQL
    'set default value of 0
    sSQL = "Update DataItemResponseHistory Set DatabaseTimeStamp = 0 where DatabaseTimeStamp is null"
    MacroADODBConnection.Execute sSQL
    
    sSQL = "ALTER Table DataItemResponseHistory ADD DatabaseTimestamp_TZ " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table MIMessage ADD MIMessageCreated_TZ " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table MIMessage ADD MIMessageSent_TZ " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table MIMessage ADD MIMessageReceived_TZ " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    'don't include this as we only use MIMessageResponseTimeStamp to join to DIR or DIRH where we can get this info anyway
'    sSQL = "ALTER Table MIMessage ADD MIMessageResponseTimestamp_TZ " & sIntegerDataType
'    MacroADODBConnection.Execute sSQL
    
    ' Lock table contains temporary records only, no need for a Timezone
    ' sSQL = "ALTER Table MacroLock ADD LockTimestamp_TZ " & sIntegerDataType
    ' MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table Message ADD MessageTimestamp_TZ " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table Message ADD MessageReceivedTimestamp_TZ " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table RSUniquenessCheck ADD CheckTimestamp_TZ " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table StudyVersion ADD VersionTimestamp_TZ " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    
    'Transfer Date limits only: Not required
    'sSQL = "ALTER Table TrialOffice ADD EffectiveFrom_TZ " & sIntegerDataType
    'MacroADODBConnection.Execute sSQL
    'sSQL = "ALTER Table TrialOffice ADD EffectiveTo_TZ " & sIntegerDataType
    'MacroADODBConnection.Execute sSQL
    
    sSQL = "ALTER Table TrialStatusHistory ADD StatusChangedTimeStamp_TZ " & sIntegerDataType
    MacroADODBConnection.Execute sSQL

    
    ' Add new columns to the NewDBcolumn table
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'TrialSubject','ImportTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'VisitInstance','ImportTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'CRFPageInstance','ImportTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'QGroupInstance','ImportTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'DataItemResponse','ResponseTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'DataItemResponse','ImportTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'DataItemResponse','DatabaseTimestamp',null,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'DataItemResponse','DatabaseTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'DataItemResponseHistory','ResponseTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'DataItemResponseHistory','ImportTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'DataItemResponseHistory','DatabaseTimestamp',null,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'DataItemResponseHistory','DatabaseTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'MIMessage','MIMessageCreated_TZ',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'MIMessage','MIMessageSent_TZ',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'MIMessage','MIMessageReceived_TZ',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'Message','MessageTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'Message','MessageReceivedTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'RSUniquenessCheck','CheckTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'StudyVersion','VersionTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,13,'TrialStatusHistory','StatusChangedTimeStamp_TZ',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL

    ' TODO: Timezone columns can contain NULL's in case data is imported from a prior version
    ' Should the import set the Timezone value to a default value, or leave NULL?

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modUpgradeDatabases.UpgradeData3_0fromto13"
    
End Sub

'---------------------------------------------------------------------
Private Sub Upgrade3_0from13to14()
'---------------------------------------------------------------------
'REM 08/10/02
'Add new column to reasonForChange table and drop all indexes and the key
'---------------------------------------------------------------------
Dim sSQL As String
Dim sIntegerDataType As String

    On Error GoTo ErrLabel
    
    'Drop the Primary Key
    sSQL = "ALTER Table ReasonForChange DROP CONSTRAINT PKReasonForChange"
    MacroADODBConnection.Execute sSQL

    'Check database type
    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.Oracle80
        sIntegerDataType = "NUMBER (6)"
    Case Else
        sIntegerDataType = "SMALLINT"
    End Select

    'add new column
    sSQL = "ALTER Table ReasonForChange ADD ReasonType " & sIntegerDataType
    MacroADODBConnection.Execute sSQL

    'add values to new columns
    sSQL = "UPDATE ReasonForChange SET ReasonType = 0 WHERE ReasonType IS NULL"
    MacroADODBConnection.Execute sSQL
    
    'Add new column to NewDBColumn table
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,14,'ReasonForChange','ReasonType',1,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modUpgradeDatabases.Upgrade3_0from13to14"
End Sub
'---------------------------------------------------------------------
Public Sub UpGradeData3_0from14to18()
'---------------------------------------------------------------------
'REM 31/10/02
'Upgade, create 2 new tables, MAVROCountry and MACROTimeZone and insert values
'Add 3 new columns to Site table, SiteLocal, SiteCountry, SiteTimeZone
'---------------------------------------------------------------------
Dim sSQL As String
Dim sText50 As String
Dim sIntegerDataType As String

    On Error GoTo ErrLabel

  '***Create Tables***
    'create 2 new tables, MACROCountry, MACROTimeZone
    Call CreateDB(eMACRODBFunction.Data, goUser.Database.DatabaseType, MacroADODBConnection, False, False, "MACROCountry", "3.0.18")
    
    Call CreateDB(eMACRODBFunction.Data, goUser.Database.DatabaseType, MacroADODBConnection, False, False, "MACROTimeZone", "3.0.18")

    ' *** Add New Tables to the table MACROTable ***
    sSQL = "INSERT INTO MACROTable VALUES ('MACROCountry','',0,0,0)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTable VALUES ('MACROTimeZone','',0,0,0)"
    MacroADODBConnection.Execute sSQL

    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.Oracle80
        sIntegerDataType = "NUMBER (6)"
        sText50 = "VARCHAR2(50)"
    Case Else
        sIntegerDataType = "SMALLINT"
        sText50 = "VARCHAR(50)"
    End Select
    
    ' *** Add New Columns ***
    'Add SiteLocale, SiteCountry, SiteTimeZone to Site table
    'Add DateTime_TZ and Loaction to LogDetails table
    sSQL = "ALTER TABLE Site ADD SiteLocale " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER TABLE Site ADD SiteCountry " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER TABLE Site ADD SiteTimeZone " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    
    sSQL = "ALTER TABLE LogDetails ADD LogDateTime_TZ " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER TABLE LogDetails ADD Location " & sText50
    MacroADODBConnection.Execute sSQL
    
    ' *** Add the new column names to the NewDBColumn table ***
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,18,'Site','SiteLocale',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,18,'Site','SiteCountry',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,18,'Site','SiteTimeZone',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,18,'LogDetails','LogDateTime_TZ',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,18,'LogDetails','Location',null,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    
    ' *** Add values to new columns ***
    'Set Location in Logdetails table to 'Local'
    sSQL = "UPDATE LogDetails SET Location = 'Local' WHERE Location IS NULL"
    MacroADODBConnection.Execute sSQL
    
    '**Insert values into the new tables***
    sSQL = "INSERT INTO MACROCountry VALUES (1026,'Bulgaria')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROCountry VALUES (1030,'Denmark')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROCountry VALUES (1031,'Germany')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROCountry VALUES (1033,'United States')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROCountry VALUES (1034,'Spain')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROCountry VALUES (1035,'Finland')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROCountry VALUES (1036,'France')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROCountry VALUES (1038,'Hungary')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROCountry VALUES (1039,'Ireland')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROCountry VALUES (1040,'Italy')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROCountry VALUES (1041,'Japan')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROCountry VALUES (1042,'Korea')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROCountry VALUES (1043,'Netherlands')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROCountry VALUES (1044,'Norway')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROCountry VALUES (1045,'Poland')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROCountry VALUES (1053,'Sweden')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROCountry VALUES (2055,'Switzerland')"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROCountry VALUES (2057,'United Kingdom')"
    MacroADODBConnection.Execute sSQL
    
    sSQL = "INSERT INTO MACROTimeZone VALUES (1,'(GMT) Greenwich Mean Time: Dublin, Edinburgh, Lisbon, London',0)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (2,'(GMT+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna',60)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (3,'(GMT+02:00) Ahtens, Istanbul, Minsk',120)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (4,'(GMT+03:00) Moscow, St.Petersburg, Volgograd',180)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (5,'(GMT+04:00) Abu Dhabi, Muscat',240)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (6,'(GMT+05:00) Islamabad, Karachi, Tashkent',300)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (7,'(GMT+06:00) Almaty, Novosibirsk',360)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (8,'(GMT+07:00) Bangkok, Hanoi, Jakarta',420)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (9,'(GMT+08:00) Beijing, Hong Kong, Singapore, Perth',480)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (10,'(GMT+09:00) Osaka, Tokyo, Seoul',540)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (11,'(GMT+10:00) Brisbane, Canberra, Melbourne, Sydney',600)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (12,'(GMT+11:00) Solomon Is, New Caledonia',660)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (13,'(GMT+12:00) Auckland, Wllington, Fiji',720)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (14,'(GMT-01:00) Cape Verde Is, Azores',-60)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (15,'(GMT-02:00) Mid Atlantic',-120)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (16,'(GMT-03:00) Greenland, Georgetown, Buenos Aires',-180)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (17,'(GMT-04:00) Aantiago, Atlantic Time (Canada)',-240)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (18,'(GMT-05:00) Eastern Time (US & Canada)',-300)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (19,'(GMT-06:00) Central Time (US & Canada), Mexico City',-360)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (20,'(GMT-07:00) Mountain Time (US & Canada), Arizona',-420)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (21,'(GMT-08:00) Pacific Time (US & Canada)',-480)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (22,'(GMT-09:00) Alaska',-540)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (23,'(GMT-10:00) Hawaii',-600)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (24,'(GMT-11:00) Midway Island, Samoa',-660)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTimeZone VALUES (25,'(GMT-12:00) Eniwetok, Kwajalein',-720)"
    MacroADODBConnection.Execute sSQL


Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modUpgradeDatabases.UpgradeData3_0from17to18"
End Sub

'----------------------------------------------------------------------------------------'
Private Function CreateDecodeSQL(bSQLServer As Boolean, sColumnName As String, _
                                    asOriginalValues As Variant, asNewValues As Variant, sElse As String) As String
'----------------------------------------------------------------------------------------'
'Create the Decode/Case SQL according to databases type
'values are sql string. ie put single quotes around a varchar datatype
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim i As Long

On Error GoTo ErrLabel

    sSQL = "CASE " & sColumnName
    
    For i = 0 To UBound(asOriginalValues)
        sSQL = sSQL & " WHEN " & asOriginalValues(i) & " THEN " & asNewValues(i)
    Next
    sSQL = sSQL & " ELSE " & sElse & " END"
    
    
    If Not bSQLServer Then
        sSQL = Replace(sSQL, "CASE ", "DECODE(")
        sSQL = Replace(sSQL, " WHEN ", ",")
        sSQL = Replace(sSQL, " THEN ", ",")
        sSQL = Replace(sSQL, " ELSE ", ",")
        sSQL = Replace(sSQL, " END", ")")
    End If
    
    CreateDecodeSQL = sSQL
Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modUpgradeDatabases.CreateDecodeSQL"
    
End Function

'---------------------------------------------------------------------
Private Sub UpGradeData2_2to3_0_7()
'---------------------------------------------------------------------
' Upgrade 2.2.x (latest) to 3.0.7 (this is first official version of MACRO 3.0)
' Changes for RQG, CRFElement captions, MIMsg status and RFC changes
' if bAsk is true each part is prompted
'---------------------------------------------------------------------
Dim sSQL As String
Dim sIntegerDataType As String
Dim sAlterColumn As String

    'set up db specific keywqords
    Select Case goUser.Database.DatabaseType
    Case MACRODatabaseType.Oracle80
        sIntegerDataType = "NUMBER (6)"
        sAlterColumn = "MODIFY"
    Case Else
        sIntegerDataType = "SMALLINT"
        sAlterColumn = "ALTER COLUMN"
    End Select

'********************************************************************************************************
'Repeating Question Group Changes

    ' *** Add New Tables to the Database ***
    Select Case goUser.Database.DatabaseType
'    Case MACRODatabaseType.Access
'
'        MacroADODBConnection.Execute "CREATE Table EFormQGroup(ClinicalTrialID Integer,VersionID SmallInt,CRFPageID Integer,QGroupID Integer,Border SmallInt,DisplayRows SmallInt,InitialRows SmallInt,MinRepeats SmallInt,MaxRepeats SmallInt, CONSTRAINT PKEFormQuestionGroup PRIMARY KEY (ClinicalTrialID,VersionID,CRFPageID,QGroupID))"
'        MacroADODBConnection.Execute "CREATE Table QGroup(ClinicalTrialID Integer,VersionID SmallInt,QGroupID Integer,QGroupCode Text(15),QGroupName Text(255),DisplayType SmallInt,CONSTRAINT PKQGroup PRIMARY KEY (ClinicalTrialID,VersionID,QGroupID))"
'        MacroADODBConnection.Execute "CREATE Table QGroupQuestion(ClinicalTrialID Integer,VersionID SmallInt,QGroupID Integer,DataItemID Integer,QOrder SmallInt,CONSTRAINT PKQGroupQuestion PRIMARY KEY (ClinicalTrialID,VersionID,QGroupID,DataItemID))"
'        MacroADODBConnection.Execute "CREATE Table QGroupInstance(ClinicalTrialID Integer,TrialSite Text(8),PersonID Integer,CRFPageTaskID Integer,QGroupID Integer,QGroupRows SmallInt,QGroupStatus SmallInt,LockStatus Byte,Changed SmallInt,ImportTimeStamp Double,CONSTRAINT PKQGroupInstance PRIMARY KEY (ClinicalTrialID,TrialSite,PersonID,CRFPageTaskID,QGroupID))"
'
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
    
        MacroADODBConnection.Execute "CREATE Table EFormQGroup(ClinicalTrialID Integer,VersionID SmallInt,CRFPageID Integer,QGroupID Integer,Border SmallInt,DisplayRows SmallInt,InitialRows SmallInt,MinRepeats SmallInt,MaxRepeats SmallInt, CONSTRAINT PKEFormQuestionGroup PRIMARY KEY (ClinicalTrialID,VersionID,CRFPageID,QGroupID))"
        MacroADODBConnection.Execute "CREATE Table QGroup(ClinicalTrialID Integer,VersionID SmallInt,QGroupID Integer,QGroupCode VarChar(15),QGroupName VarChar(255),DisplayType SmallInt,CONSTRAINT PKQGroup PRIMARY KEY (ClinicalTrialID,VersionID,QGroupID))"
        MacroADODBConnection.Execute "CREATE Table QGroupQuestion(ClinicalTrialID Integer,VersionID SmallInt,QGroupID Integer,DataItemID Integer,QOrder SmallInt,CONSTRAINT PKQGroupQuestion PRIMARY KEY (ClinicalTrialID,VersionID,QGroupID,DataItemID))"
        MacroADODBConnection.Execute "CREATE Table QGroupInstance(ClinicalTrialID Integer,TrialSite VarChar(8),PersonID Integer,CRFPageTaskID Integer,QGroupID Integer,QGroupRows SmallInt,QGroupStatus SmallInt,LockStatus TinyInt,Changed SmallInt,ImportTimeStamp Numeric(16,10),CONSTRAINT PKQGroupInstance PRIMARY KEY (ClinicalTrialID,TrialSite,PersonID,CRFPageTaskID,QGroupID))"

    Case MACRODatabaseType.Oracle80
    
        MacroADODBConnection.Execute "CREATE Table EFormQGroup(ClinicalTrialID Number(11),VersionID Number(6),CRFPageID Number(11),QGroupID Number(11),Border Number(6),DisplayRows Number(6),InitialRows Number(6),MinRepeats Number(6),MaxRepeats Number(6), CONSTRAINT PKEFormQuestionGroup PRIMARY KEY (ClinicalTrialID,VersionID,CRFPageID,QGroupID))"
        MacroADODBConnection.Execute "CREATE Table QGroup(ClinicalTrialID Number(11),VersionID Number(6),QGroupID Number(11),QGroupCode VarChar2(15),QGroupName VarChar2(255),DisplayType Number(6),CONSTRAINT PKQGroup PRIMARY KEY (ClinicalTrialID,VersionID,QGroupID))"
        MacroADODBConnection.Execute "CREATE Table QGroupQuestion(ClinicalTrialID Number(11),VersionID Number(6),QGroupID Number(11),DataItemID Number(11),QOrder Number(6),CONSTRAINT PKQGroupQuestion PRIMARY KEY (ClinicalTrialID,VersionID,QGroupID,DataItemID))"
        MacroADODBConnection.Execute "CREATE Table QGroupInstance(ClinicalTrialID Number(11),TrialSite VarChar2(8),PersonID Number(11),CRFPageTaskID Number(11),QGroupID Number(11),QGroupRows Number(6),QGroupStatus Number(6),LockStatus Number(3),Changed Number(6),ImportTimeStamp Number(16,10),CONSTRAINT PKQGroupInstance PRIMARY KEY (ClinicalTrialID,TrialSite,PersonID,CRFPageTaskID,QGroupID))"

    End Select
    

    ' *** Add New Tables to the table MACROTable ***
    sSQL = "INSERT INTO MACROTable VALUES ('EFormQGroup','210',1,0,0)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTable VALUES ('QGroupQuestion','220',1,0,0)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTable VALUES ('QGroup','230',1,0,0)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO MACROTable VALUES ('QGroupInstance','095',0,1,0)"
    MacroADODBConnection.Execute sSQL
    
    ' *** Add New Columns ***
    'Added 4 new columns to CRFElement; OwnerQGroup, QGroupId, QGroupFieldOrder, ShowStatusFlag
    'Added 1 column to DataItemResponse; RepeatNumber
    'Added 1 column to DataItemResponseHistory; RepeatNumber
    Select Case goUser.Database.DatabaseType
'    Case MACRODatabaseType.Access
'        sSQL = "ALTER Table CRFElement ADD COLUMN OwnerQGroupID INTEGER"
'        MacroADODBConnection.Execute sSQL
'        sSQL = "ALTER Table CRFElement ADD COLUMN QGroupID INTEGER"
'        MacroADODBConnection.Execute sSQL
'        sSQL = "ALTER Table CRFElement ADD COLUMN QGroupFieldOrder SMALLINT"
'        MacroADODBConnection.Execute sSQL
'        sSQL = "ALTER Table CRFElement ADD COLUMN ShowStatusFlag SMALLINT"
'        MacroADODBConnection.Execute sSQL
'        sSQL = "ALTER Table DataItemResponse ADD COLUMN RepeatNumber SMALLINT"
'        MacroADODBConnection.Execute sSQL
'        sSQL = "ALTER Table DataItemResponseHistory ADD COLUMN RepeatNumber SMALLINT"
'        MacroADODBConnection.Execute sSQL
'        sSQL = "ALTER Table MIMessage ADD COLUMN MIMessageResponseCycle SMALLINT"
'        MacroADODBConnection.Execute sSQL
        
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
        sSQL = "ALTER Table CRFElement ADD OwnerQGroupID INTEGER"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table CRFElement ADD QGroupID INTEGER"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table CRFElement ADD QGroupFieldOrder SMALLINT"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table CRFElement ADD ShowStatusFlag SMALLINT"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table DataItemResponse ADD RepeatNumber SMALLINT"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table DataItemResponseHistory ADD RepeatNumber SMALLINT"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table MIMessage ADD MIMessageResponseCycle SMALLINT"
        MacroADODBConnection.Execute sSQL
        
    Case MACRODatabaseType.Oracle80
        sSQL = "ALTER Table CRFElement ADD OwnerQGroupID NUMBER(11)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table CRFElement ADD QGroupID NUMBER(11)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table CRFElement ADD QGroupFieldOrder NUMBER(6)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table CRFElement ADD ShowStatusFlag NUMBER(6)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table DataItemResponse ADD RepeatNumber NUMBER(6)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table DataItemResponseHistory ADD RepeatNumber NUMBER(6)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table MIMessage ADD MIMessageResponseCycle NUMBER(6)"
        MacroADODBConnection.Execute sSQL
        
    End Select

    ' *** Add default values to new columns ***
    sSQL = "UPDATE CRFElement SET OwnerQGroupID = 0 WHERE OwnerQGroupID IS NULL"
    MacroADODBConnection.Execute sSQL
    sSQL = "UPDATE CRFElement SET QGroupID = 0 WHERE QGroupId IS NULL"
    MacroADODBConnection.Execute sSQL
    sSQL = "UPDATE CRFElement SET QGroupFieldOrder = 0 WHERE QGroupFieldOrder IS NULL"
    MacroADODBConnection.Execute sSQL
    sSQL = "UPDATE CRFElement SET ShowStatusFlag = 1 WHERE ShowStatusFlag IS NULL"
    MacroADODBConnection.Execute sSQL
    sSQL = "UPDATE DataItemResponse SET RepeatNumber = 1 WHERE RepeatNumber IS NULL"
    MacroADODBConnection.Execute sSQL
    sSQL = "UPDATE DataItemResponseHistory SET RepeatNumber = 1 WHERE RepeatNumber IS NULL"
    MacroADODBConnection.Execute sSQL
    sSQL = "UPDATE MIMessage SET MIMessageResponseCycle = 1 WHERE MIMessageResponseCycle IS NULL"
    MacroADODBConnection.Execute sSQL
    
    'SQLServer requires NOT NULL constraint on a column before it can be added to the Primary Key
    If (goUser.Database.DatabaseType = sqlserver) Or (goUser.Database.DatabaseType = SQLServer70) Then
        sSQL = "ALTER TABLE DataItemResponse ALTER COLUMN RepeatNumber smallint NOT NULL"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER TABLE DataItemResponseHistory ALTER COLUMN RepeatNumber smallint NOT NULL"
        MacroADODBConnection.Execute sSQL
    End If
    
    ' *** Add New Column to the Primary Key ***
    'Drop the Primary Key, have to do before adding updated one
    sSQL = "ALTER Table DataItemResponse DROP CONSTRAINT PKDataItemResponse"
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER TAble DataItemResponseHistory DROP CONSTRAINT PKDataItemResponseHistory"
    MacroADODBConnection.Execute sSQL
    'Add New Primary Key
    sSQL = "ALTER Table DataItemResponse ADD CONSTRAINT PKDataItemResponse PRIMARY KEY" & _
            "(ClinicalTrialId,TrialSite,PersonId,ResponseTaskId,RepeatNumber)"
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table DataItemResponseHistory ADD CONSTRAINT PKDataItemResponseHistory PRIMARY KEY" & _
            "(ClinicalTrialId,TrialSite,PersonId,ResponseTaskId,ResponseTimeStamp,RepeatNumber)"
    MacroADODBConnection.Execute sSQL
        
    ' *** Add the new column names to the NewDBColumn table ***
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,3,'CRFElement','OwnerQGroupID',1,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,3,'CRFElement','QGroupID',2,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,3,'CRFElement','QGroupFieldOrder',3,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,3,'CRFElement','ShowStatusFlag',4,'1','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,3,'DataItemResponse','RepeatNumber',null,'1','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,3,'DataItemResponseHistory','RepeatNumber',null,'1','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,3,'MIMessage','MIMessageResponseCycle',null,'1','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL

'********************************************************************************************************
'Caption font changes
'Added 5 new columns to CRFElement table; CaptionFontName, CaptionFontBold, CaptionFontItalic, CaptionFontSize, CaptionFontColour
    Select Case goUser.Database.DatabaseType
'    Case MACRODatabaseType.Access
'
'        sSQL = "ALTER Table CRFElement ADD COLUMN CaptionFontName Text(50)"
'        MacroADODBConnection.Execute sSQL
'        sSQL = "ALTER Table CRFElement ADD COLUMN CaptionFontBold SMALLINT"
'        MacroADODBConnection.Execute sSQL
'        sSQL = "ALTER Table CRFElement ADD COLUMN CaptionFontItalic SMALLINT"
'        MacroADODBConnection.Execute sSQL
'        sSQL = "ALTER Table CRFElement ADD COLUMN CaptionFontSize SMALLINT"
'        MacroADODBConnection.Execute sSQL
'        sSQL = "ALTER Table CRFElement ADD COLUMN CaptionFontColour INTEGER"
'        MacroADODBConnection.Execute sSQL
        
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
    
        sSQL = "ALTER Table CRFElement ADD CaptionFontName VARCHAR(50)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table CRFElement ADD CaptionFontBold SMALLINT"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table CRFElement ADD CaptionFontItalic SMALLINT"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table CRFElement ADD CaptionFontSize SMALLINT"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table CRFElement ADD CaptionFontColour INTEGER"
        MacroADODBConnection.Execute sSQL
        
    Case MACRODatabaseType.Oracle80
    
        sSQL = "ALTER Table CRFElement ADD CaptionFontName VARCHAR2(50)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table CRFElement ADD CaptionFontBold NUMBER(6)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table CRFElement ADD CaptionFontItalic NUMBER(6)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table CRFElement ADD CaptionFontSize NUMBER(6)"
        MacroADODBConnection.Execute sSQL
        sSQL = "ALTER Table CRFElement ADD CaptionFontColour NUMBER(11)"
        MacroADODBConnection.Execute sSQL
        
    End Select
    
    ' *** Add the new column names to the NewDBColumn table ***
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,4,'CRFElement','CaptionFontName',1,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,4,'CRFElement','CaptionFontBold',2,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,4,'CRFElement','CaptionFontItalic',3,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,4,'CRFElement','CaptionFontSize',4,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,4,'CRFElement','CaptionFontColour',5,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    
    'MLM 26/03/03: Copy the font attributes to caption fields.
    sSQL = "UPDATE CRFElement SET CaptionFontName = FontName, CaptionFontBold = FontBold, CaptionFontItalic = FontItalic, CaptionFontSize = FontSize, CaptionFontColour = FontColour " & _
        "WHERE DataItemId > 0 OR ControlType = 16386"
    MacroADODBConnection.Execute sSQL

'********************************************************************************************************
'New Status Icons changes
    'dataitemresponse
    sSQL = "ALTER Table DataItemResponse ADD ChangeCount " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table DataItemResponse ADD DiscrepancyStatus " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table DataItemResponse ADD SDVStatus " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table DataItemResponse ADD Notestatus " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    
    'default values for dataitemresponse
    sSQL = "Update DataItemResponse set ChangeCount = HadValue"
    MacroADODBConnection.Execute sSQL
    sSQL = "Update DataItemResponse set DiscrepancyStatus = 0"
    MacroADODBConnection.Execute sSQL
    sSQL = "Update DataItemResponse set SDVStatus = 0"
    MacroADODBConnection.Execute sSQL
    sSQL = "Update DataItemResponse set NoteStatus = 0"
    MacroADODBConnection.Execute sSQL
    
    'CRFPageInstance
    sSQL = "ALTER Table CRFPageInstance ADD DiscrepancyStatus " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table CRFPageInstance ADD SDVStatus " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table CRFPageInstance ADD Notestatus " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    
    'default values for CRFPageInstance
    sSQL = "Update CRFPageInstance set DiscrepancyStatus = 0"
    MacroADODBConnection.Execute sSQL
    sSQL = "Update CRFPageInstance set SDVStatus = 0"
    MacroADODBConnection.Execute sSQL
    sSQL = "Update CRFPageInstance set NoteStatus = 0"
    MacroADODBConnection.Execute sSQL

    
    'VisitInstance
    sSQL = "ALTER Table VisitInstance ADD DiscrepancyStatus " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table VisitInstance ADD SDVStatus " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table VisitInstance ADD Notestatus " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    
    'default values for VisitInstance
    sSQL = "Update VisitInstance set DiscrepancyStatus = 0"
    MacroADODBConnection.Execute sSQL
    sSQL = "Update VisitInstance set SDVStatus = 0"
    MacroADODBConnection.Execute sSQL
    sSQL = "Update VisitInstance set NoteStatus = 0"
    MacroADODBConnection.Execute sSQL
      
    
    'TrialSubject
    sSQL = "ALTER Table TrialSubject ADD DiscrepancyStatus " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table TrialSubject ADD SDVStatus " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER Table TrialSubject ADD Notestatus " & sIntegerDataType
    MacroADODBConnection.Execute sSQL

    'default values for TrialSubject
    sSQL = "Update TrialSubject set DiscrepancyStatus = 0"
    MacroADODBConnection.Execute sSQL
    sSQL = "Update TrialSubject set SDVStatus = 0"
    MacroADODBConnection.Execute sSQL
    sSQL = "Update TrialSubject set NoteStatus = 0"
    MacroADODBConnection.Execute sSQL

'new columns for importing from previous versions
    'dataitemresponse
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,5,'DataItemResponse','ChangeCount',1,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,5,'DataItemResponse','DiscrepancyStatus',2,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,5,'DataItemResponse','SDVStatus',3,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,5,'DataItemResponse','NoteStatus',4,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL

    'CRFPageInstance
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,5,'CRFPageInstance','DiscrepancyStatus',1,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,5,'CRFPageInstance','SDVStatus',2,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,5,'CRFPageInstance','NoteStatus',3,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    
    'VisitInstance
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,5,'VisitInstance','DiscrepancyStatus',1,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,5,'VisitInstance','SDVStatus',2,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,5,'VisitInstance','NoteStatus',3,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    
    'TrialSubject
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,5,'TrialSubject','DiscrepancyStatus',1,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,5,'TrialSubject','SDVStatus',2,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,5,'TrialSubject','NoteStatus',3,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL



'********************************************************************************************************
'Changes for CG, DataItem.Description, DataItem.MACROOnly and StudyDefinition.AREZZOUpdate
'also StudyDefiniton.RFCDefault
    sSQL = "ALTER Table DataItem ADD MACROOnly " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    'set default to 0
    sSQL = "Update DataItem SET MACROOnly = 0"
    MacroADODBConnection.Execute sSQL
    
    sSQL = "ALTER Table DataItem ADD Description VARCHAR(255)"
    MacroADODBConnection.Execute sSQL

   
    'if set to 1 MACRO SD will update AREZZO file when the opening the study
    sSQL = "ALTER Table StudyDefinition ADD AREZZOUpdateStatus " & sIntegerDataType
    MacroADODBConnection.Execute sSQL
    'set default to 0
    sSQL = "Update StudyDefinition SET AREZZOUpdateStatus = 0"
    MacroADODBConnection.Execute sSQL
    sSQL = "ALTER TABLE StudyDefinition " & sAlterColumn & " AREZZOUpdateStatus " & sIntegerDataType & " NOT NULL"
    MacroADODBConnection.Execute sSQL

'new columns for importing from previous versions
    'default value of 0 - not MACROOnly
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,7,'DataItem','MACROOnly',1,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,7,'DataItem','Description',2,'#NULL#','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL
    'default value of 0 - do not updated AREZZO file
    sSQL = "INSERT INTO NewDBColumn VALUES (3,0,7,'StudyDefinition','AREZZOUpdateStatus',2,'0','NEWCOLUMN',null)"
    MacroADODBConnection.Execute sSQL

End Sub

''---------------------------------------------------------------------
'Public Function CreateFirstDatabase() As Boolean
''---------------------------------------------------------------------
'' Check there is a database in theUpgrade a 2.2 access db to 2.2 MSDE
'' Retunrs false if unsuccessful
''---------------------------------------------------------------------
''late binding used so that changes to MACROAccess22ToMSDE22.Xfer do not require rerefernces
''switch to early binding when coding for intellisense and compilation errors
'Dim oMSDE As Object     'MACROAccess22ToMSDE22.Xfer
'Dim sLocalMachineName As String
'Dim sSAPassword As String
'Dim enFailReason As Long    'MACROAccess22ToMSDE22.eFailReason
'Dim sDBCode As String
'
'Dim sSecCon As String
'Dim sMsg As String
'Dim sDefaultPassword As String
'Dim sConnection As String
'
'
'    'default to false
'    CreateFirstDatabase = False
'
'
'    sDBCode = goUser.Database.NameOfDatabase
'    sDefaultPassword = goUser.Database.DatabasePassword
'
'    sSecCon = Connection_String(CONNECTION_MSJET_OLEDB_40, SecurityDatabasePath, , , gsSecurityDatabasePassword)
'
''late binding used so that changes to MACROAccess22ToMSDE22.Xfer do not require rerefernces
''switch to early binding when coding for intellisense and compilation errors
'    Set oMSDE = CreateObject("MACROAccess22ToMSDE22.Xfer")  'New MACROAccess22ToMSDE22.Xfer
'
'    sLocalMachineName = oMSDE.GetLocalHostName
'    sSAPassword = "" ' this may be change by CheckSAPassword
'
'    'ask them if they want to proceed
'    sMsg = "This is a first time you have used MACRO so a new database will be created now."
'    Call DialogInformation(sMsg)
'
'
''check sa password and whether MSDE is running locally
'    If Not oMSDE.CheckSAPassword(sLocalMachineName, sSAPassword, True) Then
'        DialogInformation "sa password is incorrect or an instance of MSDE is not running on this machine"
''EXIT FUNCTION
'        Exit Function
'    End If
'
'    'create the physical db
'    If Not oMSDE.CreatePhysicalDatabase(sLocalMachineName, sSAPassword, sDBCode) Then
'        DialogInformation oMSDE.GetFailReasonText(sfrFailPhysicalCreate)
''EXIT FUNCTION
'        Exit Function
'    End If
'
'    'create a new user nb user must be same as database
'    If Not oMSDE.NewUser(sLocalMachineName, sDBCode, sDefaultPassword, sSAPassword) Then
'        DialogInformation oMSDE.GetFailReasonText(sfrFailCreateUser)
''EXIT FUNCTION
'        Exit Function
'    End If
'
'    sConnection = Connection_String(CONNECTION_SQLOLEDB, sLocalMachineName, goUser.Database.NameOfDatabase, _
'                    goUser.Database.DatabaseUser, goUser.Database.DatabasePassword)
'
'    'create a fresh MACRO db
'    Call CreateDB(Data, sqlserver, sConnection, False, False)
'
'    'update the secutiry with the correct info
'    If oMSDE.UpdateSecurity(sDBCode, sSecCon, sLocalMachineName) <> sfrSuccess Then
'        DialogInformation oMSDE.GetFailReasonText(sfrFailUpdateSecurity)
'        Exit Function
'    End If
'
'     '???????? NEED TO CHECK!
'     'ensure user object uptodate
'     'goUser.Database.ServerName = sLocalMachineName
'
'    CreateFirstDatabase = True
'
'End Function

'---------------------------------------------------------------------------------
Public Sub UpgradeSecurity3_0from16to18()
'---------------------------------------------------------------------------------
'ASH 21/10/2002
'Add new table FunctionModule
'---------------------------------------------------------------------------------
Dim sSQL As String
    
    On Error GoTo ErrLabel
    
    'Create table FunctionModule
    sSQL = "CREATE Table FunctionModule(FunctionCode TEXT(10)," _
        & " MACROModule TEXT(4))"
    SecurityADODBConnection.Execute sSQL
    
    'F1001-F1008
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F1001','SM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F1002','EX')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F1003','LM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F1004','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F1005','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F1006','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F1007','DV')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F1008','QM')"
    SecurityADODBConnection.Execute (sSQL)
    'F2001-F2013
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F2001','SM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F2002','SM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F2003','SM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F2004','SM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F2005','SM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F2006','EX')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F2007','SM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F2008','SM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F2009','SM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F2010','SM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F2011','SM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F2012','SM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F2013','SM')"
    SecurityADODBConnection.Execute (sSQL)
    'F3001 - F3025
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3001','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3002','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3003','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3004','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3005','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3006','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3007','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3008','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3009','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3010','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3011','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3012','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3013','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3014','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3016','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3017','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3018','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3019','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3020','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3021','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3022','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3023','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3024','SD')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3025','SD')"
    SecurityADODBConnection.Execute (sSQL)
    'F4001 - F4008
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F4001','EX')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F4002','EX')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F4003','EX')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F4004','EX')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F4005','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F4006','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F4007','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F4008','DM')"
    SecurityADODBConnection.Execute (sSQL)
    'F5001 - F5022
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5001','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5002','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5003','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5004','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5005','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5006','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5007','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5008','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5009','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5010','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5012','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5013','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5014','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5015','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5016','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5017','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5018','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5019','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5020','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5021','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5022','DM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F5023','DM')"
    SecurityADODBConnection.Execute (sSQL)
    'F6001 - F6005
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F6001','EX')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F6002','LM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F6003','LM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F6004','LM')"
    SecurityADODBConnection.Execute (sSQL)
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F6005','LM')"
    SecurityADODBConnection.Execute (sSQL)
    
    'ASH 30/10/2002 New 'Edit question metadata description' permission
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F3026','Edit question metadata description')"
    SecurityADODBConnection.Execute sSQL

    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F3026')"
    SecurityADODBConnection.Execute sSQL
    
    sSQL = "Insert INTO FunctionModule (FunctionCode,MACROModule) "
    sSQL = sSQL & "VALUES ('F3026','SD')"
    SecurityADODBConnection.Execute (sSQL)
    
    'rem 31/10/2002
    sSQL = "ALTER TABLE LoginLog ADD LogDateTime_TZ Integer"
    SecurityADODBConnection.Execute sSQL
    sSQL = "ALTER TABLE LoginLog ADD Location TEXT(50)"
    SecurityADODBConnection.Execute sSQL

    ' *** Add values to new columns ***
    'Set Location in Logdetails table to 'Local'
    sSQL = "UPDATE LoginLog SET Location = 'Local' WHERE Location IS NULL"
    SecurityADODBConnection.Execute sSQL


Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modUpgradeDatabases.UpGradeSecurity3_0from16to17"
End Sub

'---------------------------------------------------------------------------------
Private Sub UpgradeSecurity3_0from18to20()
'---------------------------------------------------------------------------------
'REM 05/11/02
'Encrypt existing database passwords and database users
'---------------------------------------------------------------------------------
Dim sSQL As String
Dim rsPswdUser As ADODB.Recordset
Dim sDatabasePswd As String
Dim sDatabaseUser As String
Dim sEncryptedDatabasePswd As String
Dim sEncryptedDatabaseUser As String
Dim sDatabaseCode As String

    On Error GoTo ErrLabel
    
    'Change the DatabasePassword and databaseUser field size in the Databases table from text50 to text100 to hold new hashed password
    sSQL = "ALTER Table Databases ALTER Column DatabasePassword TEXT(128)"
    SecurityADODBConnection.Execute sSQL
    sSQL = "ALTER Table Databases ALTER Column DatabaseUser TEXT(128)"
    SecurityADODBConnection.Execute sSQL
    
    'get all existing database passswords and Database users
    sSQL = "SELECT DatabaseCode, DatabasePassword, DatabaseUser" _
         & " FROM Databases"
    Set rsPswdUser = New ADODB.Recordset
    rsPswdUser.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    rsPswdUser.MoveFirst
    'loop through them and encrypt
    Do While Not rsPswdUser.EOF
        sDatabaseCode = rsPswdUser!DatabaseCode
        sDatabasePswd = RemoveNull(rsPswdUser!DatabasePassword)
        sDatabaseUser = RemoveNull(rsPswdUser!DatabaseUser)
        
        If sDatabasePswd = "" Then
            sEncryptedDatabasePswd = "null"
        Else
            sEncryptedDatabasePswd = "'" & EncryptString(sDatabasePswd) & "'"
        End If
        
        If sDatabaseUser = "" Then
            sEncryptedDatabaseUser = "null"
        Else
            sEncryptedDatabaseUser = "'" & EncryptString(sDatabaseUser) & "'"
        End If
        
        sSQL = "UPDATE Databases SET DatabasePassword = " & sEncryptedDatabasePswd & ", " _
            & " DatabaseUser = " & sEncryptedDatabaseUser _
            & " WHERE DatabaseCode = '" & sDatabaseCode & "'"
        SecurityADODBConnection.Execute sSQL
    
        rsPswdUser.MoveNext
    Loop

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modUpgradeDatabases.UpgradeSecurity3_0from18to20"
End Sub


